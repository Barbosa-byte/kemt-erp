import sys
import os
import sqlite3
import shutil
import bcrypt
import openpyxl
from datetime import datetime, timedelta
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as ExcelImage
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QStackedWidget, QTabWidget,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView,
    QFormLayout, QFileDialog, QGroupBox, QDialog, QMenu, QStatusBar
)

APP_NAME = "Kemt"
COPYRIGHT_TEXT = "© Developed by: KO (01023368006) - All Rights Reserved"

# -------------------------------------------------------------
# قواميس اللغات والتلميحات التوضيحية
# -------------------------------------------------------------
TRANSLATIONS = {
    'en': {
        'app_title': f'{APP_NAME} - Import & Shipping Management ERP',
        'file_btn': 'File ▾',
        'entities_btn': 'Entities Directory',
        'users_btn': 'User Management',
        'logs_btn': 'Audit Logs (Activity Tracking)',
        'about_btn': 'About Kemt',
        'nav_pur': 'Purchases',
        'nav_ship': 'Shipping & Containers',
        'nav_pay': 'Payments & Expenses',
        'nav_rep': 'Financial Summary',
        'nav_cnt_rep': 'Container Manifest Report',
        'nav_client_dossier': 'Client 360° Dossier',
        'fiscal_year_lbl': 'Fiscal Year:',
        'all_years': 'All Fiscal Years (All-Time)',
        'filter_btn': 'Filter',
        'inv_box': 'Commercial Invoice Details (Client Purchases)',
        'inv_num': 'Invoice No:',
        'client': 'Client:',
        'supplier': 'Supplier (Factory):',
        'agent': 'Shipping Agent:',
        'shipping_line': 'Shipping Line:',
        'inv_total': 'Total Amount:',
        'inv_net': 'Net Amount (EGP):',
        'inv_cbm': 'Total CBM:',
        'save_inv_btn': 'Save Commercial Invoice',
        'update_inv_btn': 'Update Purchase Invoice',
        'cancel_edit_btn': 'Cancel Edit',
        'cnt_box': 'Container Logistics & Landed Costs',
        'cnt_num': 'Container No:',
        'bol_num': 'B/L No:',
        'sea_freight': 'Sea Freight:',
        'customs': 'Customs & Taxes (EGP):',
        'commission': 'Commission / Handling (EGP):',
        'total_cnt_cost': 'Total Container Cost (EGP):',
        'save_cnt_btn': 'Save Container Details',
        'quick_track_box': 'Quick Status Tracking (Data Entry & Admin)',
        'btn_update_status_user': 'Update Status',
        'subtab_client_pay': 'Client Payments',
        'subtab_expenses': 'Clearance & Expenses',
        'amount': 'Amount:',
        'pay_method': 'Payment Method:',
        'attach_doc': 'Attach Receipt / Slip',
        'save_pay_btn': 'Save Client Payment',
        'target_cnt': 'Target Container:',
        'exp_category': 'Expense Item:',
        'exp_amount': 'Expense Amount:',
        'notes': 'Notes / Description:',
        'save_exp_btn': 'Save Container Expense',
        'export_excel_btn': 'Export Financial Summary to Excel',
        'export_cnt_excel_btn': 'Export Container Manifest to Excel',
        'export_client_excel_btn': 'Export Client Dossier to Excel',
        
        # Tooltips - English
        'tip_pur_inv_num': 'Unique commercial invoice number issued by the factory.',
        'tip_pur_client': 'Select client who placed the purchasing order.',
        'tip_pur_supplier': 'Select foreign supplier/manufacturer in China.',
        'tip_pur_curr': 'Original billing currency (USD, RMB, EUR, EGP).',
        'tip_pur_fx': 'Exchange rate to convert foreign currency into local Egyptian Pounds (EGP).',
        'tip_pur_total': 'Total invoice value in original foreign currency.',
        'tip_pur_net': 'Automatically calculated local equivalent amount in EGP (Total * FX).',
        'tip_pur_cbm': 'Total cubic meters (CBM) of the purchased items.',
        'tip_pur_import_excel': 'Upload factory commercial invoice or packing list directly from Excel.',
        'tip_pur_save_btn': 'Store invoice and cargo items into system records.',
        
        'tip_cnt_num': 'Unique container tracking number (e.g., MSKU1234567).',
        'tip_cnt_client': 'Primary client owning this container.',
        'tip_cnt_agent': 'Freight forwarding company or shipping agent.',
        'tip_cnt_line': 'Ocean carrier name (Maersk, MSC, COSCO, etc.).',
        'tip_cnt_bol': 'Ocean Bill of Lading (B/L) number.',
        'tip_cnt_status': 'Current transit stage of the container.',
        'tip_cnt_freight': 'Ocean freight cost in foreign currency.',
        'tip_cnt_freight_fx': 'Exchange rate used to convert freight cost into EGP.',
        'tip_cnt_customs': 'Customs clearance duties, taxes and VAT paid in Egypt in EGP.',
        'tip_cnt_comm': 'Clearance agency commission, truck haulage, and handling fees in EGP.',
        'tip_cnt_total_egp': 'Total landed trip costs in EGP, allocated across cargo items.',
        'tip_cnt_quick_status': 'Quickly update shipping stage without opening management reports.',
        
        'tip_pay_client': 'Select client submitting this financial payment.',
        'tip_pay_amt': 'Paid amount value.',
        'tip_pay_fx': 'Exchange rate applied to convert payment to EGP.',
        'tip_pay_meth': 'Method of payment transaction (Cash, Bank SWIFT, etc.).',
        'tip_pay_attach': 'Attach deposit slip, bank transfer screenshot, or receipt scan.',
        'tip_exp_cnt': 'Target container bearing this clearance or demurrage expense.',
        'tip_exp_cat': 'Category of unexpected or port-related expenses.',
        'tip_exp_amt': 'Expense amount paid locally on behalf of client.'
    },
    'ar': {
        'app_title': f'{APP_NAME} - نظام إدارة الاستيراد والشحن',
        'file_btn': 'ملف ▾',
        'entities_btn': 'دليل العملاء والموردين',
        'users_btn': 'إدارة المستخدمين',
        'logs_btn': 'سجل رقابة العمليات (Audit Logs)',
        'about_btn': f'عن برنامج {APP_NAME}',
        'nav_pur': 'المشتريات',
        'nav_ship': 'شحن وإستلام حاوية',
        'nav_pay': 'التسديدات والمصروفات',
        'nav_rep': 'الملخص المالي',
        'nav_cnt_rep': 'تقرير تفاصيل الحاوية ومشمولها',
        'nav_client_dossier': 'الملف الشامل للعميل (360°)',
        'fiscal_year_lbl': 'السنة المالية:',
        'all_years': 'كل السنوات المالية (إجمالي)',
        'filter_btn': 'تصفية',
        'inv_box': 'بيانات الفاتورة التجارية (مشتريات عميل من المصنع)',
        'inv_num': 'رقم الفاتورة:',
        'client': 'العميل:',
        'supplier': 'المورد:',
        'agent': 'وكيل الشحن:',
        'shipping_line': 'الخط الملاحي:',
        'inv_total': 'إجمالي الفاتورة:',
        'inv_net': 'صافي الفاتورة (بالجنيه EGP):',
        'inv_cbm': 'إجمالي CBM:',
        'save_inv_btn': 'حفظ فاتورة المشتريات',
        'update_inv_btn': 'تحديث بيانات الفاتورة',
        'cancel_edit_btn': 'إلغاء التعديل',
        'cnt_box': 'بيانات الحاوية وتكاليف الشحن والتخليص',
        'cnt_num': 'رقم الحاوية:',
        'bol_num': 'رقم البوليصة:',
        'sea_freight': 'النولون البحري:',
        'customs': 'الجمارك والضرائب (EGP):',
        'commission': 'العمولة والتخليص (EGP):',
        'total_cnt_cost': 'إجمالي تكلفة الحاوية (بالجنيه EGP):',
        'save_cnt_btn': 'تسجيل وحفظ بيانات الحاوية',
        'quick_track_box': 'متابعة وتحديث حالة الحاوية (متاح لمدخل البيانات والأدمن)',
        'btn_update_status_user': 'تحديث وحفظ الحالة',
        'subtab_client_pay': 'دفعات العملاء',
        'subtab_expenses': 'المصروفات والتخليص',
        'amount': 'المبلغ:',
        'pay_method': 'طريقة الدفع:',
        'attach_doc': 'إرفاق إيصال التحويل / السداد',
        'save_pay_btn': 'حفظ دفعة العميل',
        'target_cnt': 'الحاوية المعنية:',
        'exp_category': 'بند المصروف:',
        'exp_amount': 'المبلغ المنصرف:',
        'notes': 'بيان وملاحظات:',
        'save_exp_btn': 'تسجيل المصروف على الحاوية',
        'export_excel_btn': 'تصدير الملخص المالي إلى Excel',
        'export_cnt_excel_btn': 'تصدير مشمول الحاوية إلى Excel',
        'export_client_excel_btn': 'تصدير ملف العميل الشامل إلى Excel',
        
        # Tooltips - Arabic
        'tip_pur_inv_num': 'رقم الفاتورة التجارية الصادرة من المصنع أو المورد.',
        'tip_pur_client': 'اختر العميل صاحب طلبية الشراء.',
        'tip_pur_supplier': 'المصنع أو المورد الصيني/الأجنبي القائم بالتصدير.',
        'tip_pur_curr': 'عملة فاتورة الشراء الأصلية (دولار، يوان، يورو، جنيه).',
        'tip_pur_fx': 'سعر الصرف المعتمد لتحويل الفاتورة إلى الجنيه المصري.',
        'tip_pur_total': 'إجمالي قيمة البضاعة بالعملة الأجنبية الأصلية.',
        'tip_pur_net': 'المعادل بالجنيه المصري (يُحسب آلياً = الإجمالي × سعر الصرف).',
        'tip_pur_cbm': 'الحجم الكلي للشحنة بالمتر المكعب (CBM).',
        'tip_pur_import_excel': 'استيراد جدول الأصناف بالكامل مباشرة من ملف إكسيل.',
        'tip_pur_save_btn': 'حفظ بيانات الفاتورة والأصناف في قاعدة البيانات.',
        
        'tip_cnt_num': 'الرقم التعريفي للحاوية (مثل MSKU1234567).',
        'tip_cnt_client': 'العميل صاحب البضاعة المشحونة بهذه الحاوية.',
        'tip_cnt_agent': 'شركة الملاحة أو وكيل الشحن المسؤول عن الحجز.',
        'tip_cnt_line': 'اسم الخط الملاحي الناقل (Maersk, MSC, COSCO...).',
        'tip_cnt_bol': 'رقم بوليصة الشحن البحرية (Bill of Lading).',
        'tip_cnt_status': 'المرحلة اللوجستية الحالية للحاوية.',
        'tip_cnt_freight': 'تكلفة النولون البحري بالعملة الأجنبية.',
        'tip_cnt_freight_fx': 'سعر صرف عملة النولون للتحويل إلى الجنيه المصري.',
        'tip_cnt_customs': 'إجمالي الرسوم الجمركية وضريبة القيمة المضافة المسددة في مصر (بالجنيه).',
        'tip_cnt_comm': 'أتعاب التخليص الجمركي ومصروفات النقل الداخلي (بالجنيه).',
        'tip_cnt_total_egp': 'إجمالي مصاريف الحاوية بالجنيه، ويتم توزيعها على الأصناف لحساب تكلفة القطعة.',
        'tip_cnt_quick_status': 'تحديث مسار الحاوية فوراً دون الحاجة لفتح شاشات التقارير.',
        
        'tip_pay_client': 'اختر العميل الذي قام بسداد الدفعة النقدية.',
        'tip_pay_amt': 'المبلغ المسدد من العميل.',
        'tip_pay_fx': 'سعر الصرف لتحويل الدفعة إلى الجنيه المصري.',
        'tip_pay_meth': 'طريقة السداد (نقدي، تحويل بنكي، شيك...).',
        'tip_pay_attach': 'إرفاق صورة إشعار التحويل البنكي أو إيصال الإيداع.',
        'tip_exp_cnt': 'الحاوية التي تم الصرف عليها أثناء وجودها في الميناء.',
        'tip_exp_cat': 'بند المصروف (أرضيات، غرامات، فحص ومعامل، تعتيق...).',
        'tip_exp_amt': 'المبلغ المنصرف من خزينتك ويُقيد كمديونية على العميل.'
    },
    'zh': {
        'app_title': f'{APP_NAME} - 进出口与清关货运管理系统',
        'file_btn': '文件 ▾',
        'entities_btn': '客户与供应商名录',
        'users_btn': '用户与权限管理',
        'logs_btn': '审计日志 (操作记录)',
        'about_btn': f'关于 {APP_NAME}',
        'nav_pur': '采购发票',
        'nav_ship': '集装箱与海运',
        'nav_pay': '付款与各项费用',
        'nav_rep': '财务汇总报表',
        'nav_cnt_rep': '集装箱货品清单报告',
        'nav_client_dossier': '客户全方位档案(360°)',
        'fiscal_year_lbl': '会计年度:',
        'all_years': '所有会计年度 (全部数据)',
        'filter_btn': '筛选',
        'inv_box': '客户采购商业发票明细',
        'inv_num': '发票号码:',
        'client': '客户:',
        'supplier': '供应商(工厂):',
        'agent': '货运代理:',
        'shipping_line': '船运公司:',
        'inv_total': '发票总额:',
        'inv_net': '折合本币金额(EGP):',
        'inv_cbm': '总体积(CBM):',
        'save_inv_btn': '保存商业发票',
        'update_inv_btn': '更新采购发票',
        'cancel_edit_btn': '取消编辑',
        'cnt_box': '集装箱海运与报关明细',
        'cnt_num': '集装箱号(柜号):',
        'bol_num': '提单号码(B/L):',
        'sea_freight': '海运费:',
        'customs': '海关关税(EGP):',
        'commission': '代理佣金(EGP):',
        'total_cnt_cost': '集装箱总费用(EGP):',
        'save_cnt_btn': '保存货柜与费用',
        'quick_track_box': '集装箱状态追踪更新',
        'btn_update_status_user': '更新状态',
        'subtab_client_pay': '客户付款',
        'subtab_expenses': '清关与港口杂费',
        'amount': '金额:',
        'pay_method': '付款方式:',
        'attach_doc': '上传转账单据/凭证',
        'save_pay_btn': '保存客户付款',
        'target_cnt': '对应集装箱:',
        'exp_category': '费用科目:',
        'exp_amount': '支出金额:',
        'notes': '备注说明:',
        'save_exp_btn': '保存集装箱费用',
        'export_excel_btn': '导出财务汇总至Excel',
        'export_cnt_excel_btn': '导出货柜清单至Excel',
        'export_client_excel_btn': '导出客户档案至Excel',
        
        # Tooltips - Chinese
        'tip_pur_inv_num': '工厂开具的商业采购发票号码。',
        'tip_pur_client': '选择此采购批次的埃及委托客户。',
        'tip_pur_supplier': '中国供货厂家或出口商企业名称。',
        'tip_pur_curr': '发票结算外币种类 (USD, RMB, EUR, EGP)。',
        'tip_pur_fx': '折算为埃及镑 (EGP) 的当前记账汇率。',
        'tip_pur_total': '原始币种发票票面总金额。',
        'tip_pur_net': '系统自动计算的折合埃及镑总值 (总额 × 汇率)。',
        'tip_pur_cbm': '采购商品货物总体积 (立方米 CBM)。',
        'tip_pur_import_excel': '从Excel装箱单快速批量导入货物明细。',
        'tip_pur_save_btn': '保存发票明细并更新采购数据库。',
        
        'tip_cnt_num': '国际集装箱箱号 (例如 MSKU1234567)。',
        'tip_cnt_client': '负责清关接收的埃及主客户。',
        'tip_cnt_agent': '负责订舱与操作的海运货代公司。',
        'tip_cnt_line': '实际承运船运公司 (马士基、中远海运等)。',
        'tip_cnt_bol': '海运提单号码 (B/L Number)。',
        'tip_cnt_status': '集装箱当前实际运输与清关流转阶段。',
        'tip_cnt_freight': '外币海运费总金额。',
        'tip_cnt_freight_fx': '用于将外币海运费折算为埃及镑的汇率。',
        'tip_cnt_customs': '在埃及海关实际缴纳的关税和增值税 (EGP)。',
        'tip_cnt_comm': '当地清关代理费与内陆集卡拖车运杂费 (EGP)。',
        'tip_cnt_total_egp': '货柜总落地成本 (EGP)，按CBM分摊至各单品。',
        'tip_cnt_quick_status': '快速流转货柜运输状态，无需进入财务报表。',
        
        'tip_pay_client': '选择付款记账的埃及客户。',
        'tip_pay_amt': '客户实际支付的收款金额。',
        'tip_pay_fx': '折算埃及镑结算汇率。',
        'tip_pay_meth': '结算支付途径 (现金、银行电汇SWIFT等)。',
        'tip_pay_attach': '上传银行水单或付款凭条图片存证。',
        'tip_exp_cnt': '港口清关或滞箱费用所属目标集装箱。',
        'tip_exp_cat': '港口杂费明细 (仓储费、滞箱费、查验费等)。',
        'tip_exp_amt': '代垫付支出的本地费用，记入客户应付账款。'
    }
}

# -------------------------------------------------------------
# قاعدة البيانات وتنظيف الـ Logs تلقائياً
# -------------------------------------------------------------
def cleanup_old_audit_logs(days_threshold=40):
    try:
        conn = sqlite3.connect("import_enterprise.db")
        cutoff_date = (datetime.now() - timedelta(days=days_threshold)).strftime("%Y-%m-%d %H:%M:%S")
        conn.execute("DELETE FROM audit_logs WHERE timestamp < ?", (cutoff_date,))
        conn.commit()
        conn.close()
    except Exception:
        pass

def init_database():
    if not os.path.exists("attachments"):
        os.makedirs("attachments")

    conn = sqlite3.connect("import_enterprise.db")
    c = conn.cursor()

    c.execute("""CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS entities (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        entity_type TEXT NOT NULL,
        phone TEXT
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS client_invoices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_num TEXT UNIQUE NOT NULL,
        client_id INTEGER,
        supplier_id INTEGER,
        currency TEXT DEFAULT 'USD',
        exchange_rate REAL DEFAULT 1.0,
        total_amount REAL DEFAULT 0,
        total_amount_local REAL DEFAULT 0,
        net_amount REAL DEFAULT 0,
        total_cbm REAL DEFAULT 0,
        fiscal_year INTEGER
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS invoice_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_id INTEGER NOT NULL,
        item_code TEXT,
        item_ar TEXT,
        item_cn TEXT,
        cartons INTEGER DEFAULT 0,
        pcs_per_carton INTEGER DEFAULT 1,
        total_pcs INTEGER DEFAULT 0,
        piece_price REAL DEFAULT 0,
        total_price REAL DEFAULT 0,
        cbm_per_carton REAL DEFAULT 0,
        total_cbm REAL DEFAULT 0,
        cbm REAL DEFAULT 0,
        weight REAL DEFAULT 0,
        shipped_status INTEGER DEFAULT 0,
        FOREIGN KEY (invoice_id) REFERENCES client_invoices(id)
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS containers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        container_num TEXT UNIQUE NOT NULL,
        client_id INTEGER,
        agent_id INTEGER,
        shipping_line TEXT,
        bill_of_lading TEXT,
        freight_currency TEXT DEFAULT 'USD',
        sea_freight REAL DEFAULT 0,
        freight_exchange_rate REAL DEFAULT 50.0,
        customs_cost REAL DEFAULT 0,
        commission REAL DEFAULT 0,
        total_container_cost REAL DEFAULT 0,
        status TEXT DEFAULT 'قيد الشحن (In-Transit)',
        arrival_date TEXT,
        currency TEXT DEFAULT 'EGP',
        exchange_rate REAL DEFAULT 1.0,
        stocked_status INTEGER DEFAULT 0,
        fiscal_year INTEGER
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS container_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        container_id INTEGER NOT NULL,
        invoice_item_id INTEGER,
        source_invoice_num TEXT,
        item_code TEXT,
        item_ar TEXT,
        item_cn TEXT,
        cartons INTEGER DEFAULT 0,
        pcs_per_carton INTEGER DEFAULT 1,
        total_pcs INTEGER DEFAULT 0,
        piece_price REAL DEFAULT 0,
        total_price REAL DEFAULT 0,
        cbm REAL DEFAULT 0,
        total_cbm REAL DEFAULT 0,
        weight REAL DEFAULT 0,
        allocated_expense REAL DEFAULT 0,
        landed_cost_unit REAL DEFAULT 0,
        FOREIGN KEY (container_id) REFERENCES containers(id)
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS ledger (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tx_category TEXT NOT NULL,
        entity_id INTEGER,
        container_id INTEGER,
        currency TEXT DEFAULT 'EGP',
        exchange_rate REAL DEFAULT 1.0,
        amount REAL NOT NULL,
        amount_local REAL NOT NULL,
        tx_date TEXT,
        method TEXT,
        doc_attachment TEXT,
        notes TEXT,
        fiscal_year INTEGER,
        created_by TEXT
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS supplier_ledger (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        supplier_id INTEGER NOT NULL,
        invoice_id INTEGER,
        tx_type TEXT NOT NULL,
        amount REAL NOT NULL DEFAULT 0,
        amount_local REAL NOT NULL DEFAULT 0,
        currency TEXT DEFAULT 'EGP',
        exchange_rate REAL DEFAULT 1.0,
        supplier_invoice_num TEXT,
        reason TEXT NOT NULL,
        tx_date TEXT,
        method TEXT,
        created_by TEXT,
        FOREIGN KEY (supplier_id) REFERENCES entities(id),
        FOREIGN KEY (invoice_id) REFERENCES client_invoices(id)
    )""")
    # Migration: older databases may already have supplier_ledger without the
    # supplier_invoice_num column. CREATE TABLE IF NOT EXISTS does not alter
    # an existing SQLite table, so add the column before creating the index.
    c.execute("PRAGMA table_info(supplier_ledger)")
    supplier_ledger_columns = {row[1] for row in c.fetchall()}

    # Migration for older supplier_ledger tables. SQLite's
    # CREATE TABLE IF NOT EXISTS does not modify an existing table.
    supplier_ledger_migrations = {
        "invoice_id": "INTEGER",
        "currency": "TEXT DEFAULT 'EGP'",
        "exchange_rate": "REAL DEFAULT 1.0",
        "supplier_invoice_num": "TEXT",
        "method": "TEXT",
        "created_by": "TEXT",
    }
    for col_name, col_definition in supplier_ledger_migrations.items():
        if col_name not in supplier_ledger_columns:
            c.execute(
                f"ALTER TABLE supplier_ledger ADD COLUMN {col_name} {col_definition}"
            )

    c.execute("""CREATE UNIQUE INDEX IF NOT EXISTS ux_supplier_invoice_debt
                 ON supplier_ledger(supplier_id, supplier_invoice_num, tx_type)
                 WHERE tx_type = 'INVOICE' AND supplier_invoice_num IS NOT NULL AND supplier_invoice_num <> ''""")

    c.execute("""CREATE TABLE IF NOT EXISTS fiscal_locks (
        year INTEGER PRIMARY KEY,
        is_locked INTEGER DEFAULT 0,
        locked_by TEXT,
        locked_at TEXT
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS audit_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TEXT NOT NULL,
        username TEXT NOT NULL,
        action TEXT NOT NULL,
        details TEXT NOT NULL
    )""")

    tables_alterations = {
        "client_invoices": [("currency", "TEXT DEFAULT 'USD'"), ("exchange_rate", "REAL DEFAULT 1.0"), ("total_amount_local", "REAL DEFAULT 0")],
        "containers": [
            ("freight_currency", "TEXT DEFAULT 'USD'"),
            ("freight_exchange_rate", "REAL DEFAULT 50.0"),
            ("currency", "TEXT DEFAULT 'EGP'"),
            ("exchange_rate", "REAL DEFAULT 1.0"),
            ("stocked_status", "INTEGER DEFAULT 0")
        ],
        "container_items": [
            ("allocated_expense", "REAL DEFAULT 0"), ("landed_cost_unit", "REAL DEFAULT 0"),
            ("invoice_item_id", "INTEGER"), ("source_invoice_num", "TEXT"),
            ("pcs_per_carton", "INTEGER DEFAULT 1"), ("total_pcs", "INTEGER DEFAULT 0"),
            ("total_price", "REAL DEFAULT 0"), ("total_cbm", "REAL DEFAULT 0")
        ],
        "invoice_items": [
            ("pcs_per_carton", "INTEGER DEFAULT 1"), ("total_pcs", "INTEGER DEFAULT 0"),
            ("total_price", "REAL DEFAULT 0"), ("cbm_per_carton", "REAL DEFAULT 0"),
            ("total_cbm", "REAL DEFAULT 0"), ("shipped_status", "INTEGER DEFAULT 0")
        ],
        "ledger": [("currency", "TEXT DEFAULT 'EGP'"), ("exchange_rate", "REAL DEFAULT 1.0"), ("amount_local", "REAL NOT NULL DEFAULT 0")]
    }

    for tbl, cols in tables_alterations.items():
        c.execute(f"PRAGMA table_info({tbl})")
        existing = [r[1] for r in c.fetchall()]
        for col_name, col_def in cols:
            if col_name not in existing:
                try:
                    c.execute(f"ALTER TABLE {tbl} ADD COLUMN {col_name} {col_def}")
                except Exception:
                    pass

    c.execute("SELECT COUNT(*) FROM users")
    if c.fetchone()[0] == 0:
        p_admin = bcrypt.hashpw(b"admin123", bcrypt.gensalt()).decode()
        p_user = bcrypt.hashpw(b"user123", bcrypt.gensalt()).decode()
        c.execute("INSERT INTO users VALUES (NULL, 'admin', ?, 'admin')", (p_admin,))
        c.execute("INSERT INTO users VALUES (NULL, 'user', ?, 'data_entry')", (p_user,))
        conn.commit()

    conn.close()
    cleanup_old_audit_logs(40)

def record_log(username, action, details):
    try:
        conn = sqlite3.connect("import_enterprise.db")
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute("INSERT INTO audit_logs (timestamp, username, action, details) VALUES (?, ?, ?, ?)",
                     (now_str, username, action, details))
        conn.commit()
        conn.close()
    except Exception:
        pass

def is_year_locked(year):
    conn = sqlite3.connect("import_enterprise.db")
    c = conn.cursor()
    c.execute("SELECT is_locked FROM fiscal_locks WHERE year = ?", (year,))
    row = c.fetchone()
    conn.close()
    return bool(row and row[0] == 1)

# -------------------------------------------------------------
# نافذة سجل رقابة العمليات وحذف السجلات
# -------------------------------------------------------------
class AuditLogDialog(QDialog):
    def __init__(self, current_user):
        super().__init__()
        self.current_user = current_user
        self.setWindowTitle("System Audit Logs - سجل رقابة العمليات وحركات الموظفين")
        self.resize(880, 530)
        layout = QVBoxLayout()

        top_bar = QHBoxLayout()
        top_bar.addWidget(QLabel("<b>Audit Trail (متابعة العمليات المنفذة - تنظيف تلقائي للأقدم من 40 يوماً):</b>"))
        top_bar.addStretch()

        self.btn_refresh = QPushButton("🔄 Refresh / تحديث")
        self.btn_refresh.setStyleSheet("background-color: #0d6efd; color: white; font-weight: bold; padding: 5px 12px;")
        self.btn_refresh.clicked.connect(self.load_logs)
        top_bar.addWidget(self.btn_refresh)

        self.btn_clear_all = QPushButton("🗑️ Clear All Logs / مسح السجل يدوياً")
        self.btn_clear_all.setStyleSheet("background-color: #dc3545; color: white; font-weight: bold; padding: 5px 12px;")
        self.btn_clear_all.clicked.connect(self.manual_clear_logs)
        top_bar.addWidget(self.btn_clear_all)

        layout.addLayout(top_bar)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Timestamp (التاريخ والوقت)", "User (المستخدم)", "Action (نوع الإجراء)", "Details (التفاصيل)"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        layout.addWidget(self.table)

        self.setLayout(layout)
        self.load_logs()

    def load_logs(self):
        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        c.execute("SELECT timestamp, username, action, details FROM audit_logs ORDER BY id DESC LIMIT 500")
        rows = c.fetchall()
        conn.close()

        self.table.setRowCount(0)
        for idx, (t, u, a, d) in enumerate(rows):
            self.table.insertRow(idx)
            self.table.setItem(idx, 0, QTableWidgetItem(t))
            user_item = QTableWidgetItem(u)
            user_item.setForeground(Qt.blue if u != 'admin' else Qt.darkRed)
            self.table.setItem(idx, 1, user_item)
            self.table.setItem(idx, 2, QTableWidgetItem(a))
            self.table.setItem(idx, 3, QTableWidgetItem(d))

    def manual_clear_logs(self):
        reply = QMessageBox.question(
            self, "Confirm Purge",
            "Are you sure you want to completely clear the activity audit logs?\nهل أنت متأكد من رغبتك في تصفير ومسح سجل الحركات بالكامل؟",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            conn = sqlite3.connect("import_enterprise.db")
            conn.execute("DELETE FROM audit_logs")
            conn.commit()
            conn.close()
            record_log(self.current_user, "CLEAR_LOGS", "Manually cleared all system audit logs")
            self.load_logs()
            QMessageBox.information(self, "Logs Cleared", "Audit logs cleared successfully.")

# -------------------------------------------------------------
# نافذة قفل السنة المالية
# -------------------------------------------------------------
class FiscalLockDialog(QDialog):
    def __init__(self, current_user):
        super().__init__()
        self.current_user = current_user
        self.setWindowTitle("Fiscal Year Lock Management - إقفال السنوات المالية")
        self.resize(460, 360)
        layout = QVBoxLayout()

        layout.addWidget(QLabel("<b>Manage Fiscal Years Status (حظر وتجميد السنوات المالية المغلقة):</b>"))
        
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Fiscal Year", "Status", "Action"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        self.setLayout(layout)
        self.load_years()

    def load_years(self):
        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        c.execute("SELECT DISTINCT fiscal_year FROM ledger WHERE fiscal_year IS NOT NULL")
        years = [r[0] for r in c.fetchall()]
        curr_y = datetime.now().year
        if curr_y not in years:
            years.append(curr_y)
        years = sorted(list(set(years)), reverse=True)

        self.table.setRowCount(0)
        for idx, y in enumerate(years):
            c.execute("SELECT is_locked FROM fiscal_locks WHERE year = ?", (y,))
            r = c.fetchone()
            locked = bool(r and r[0] == 1)

            self.table.insertRow(idx)
            self.table.setItem(idx, 0, QTableWidgetItem(f"FY {y}"))
            status_item = QTableWidgetItem("🔒 LOCKED (مغلقة)" if locked else "🔓 OPEN (متاحة)")
            status_item.setForeground(Qt.red if locked else Qt.darkGreen)
            self.table.setItem(idx, 1, status_item)

            btn = QPushButton("Unlock Year" if locked else "Lock Year")
            btn.setStyleSheet("background-color: #ffc107;" if locked else "background-color: #dc3545; color: white;")
            btn.clicked.connect(lambda checked, yr=y, lk=locked: self.toggle_lock(yr, lk))
            self.table.setCellWidget(idx, 2, btn)
        conn.close()

    def toggle_lock(self, year, current_state):
        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        new_state = 0 if current_state else 1
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        c.execute("""INSERT INTO fiscal_locks (year, is_locked, locked_by, locked_at) 
                     VALUES (?, ?, ?, ?)
                     ON CONFLICT(year) DO UPDATE SET is_locked=?, locked_by=?, locked_at=?""",
                  (year, new_state, self.current_user, now_str, new_state, self.current_user, now_str))
        conn.commit()
        conn.close()
        record_log(self.current_user, "FISCAL_LOCK_TOGGLE", f"{'Unlocked' if current_state else 'Locked'} Fiscal Year {year}")
        self.load_years()

# -------------------------------------------------------------
# نافذة إدارة المستخدمين وتغيير كلمة المرور
# -------------------------------------------------------------
class UserManagerDialog(QDialog):
    def __init__(self, current_user):
        super().__init__()
        self.current_user = current_user
        self.setWindowTitle(f"{APP_NAME} - User Management & Password Reset")
        self.resize(580, 420)
        layout = QVBoxLayout()

        form = QFormLayout()
        self.u_input = QLineEdit()
        self.p_input = QLineEdit()
        self.p_input.setEchoMode(QLineEdit.Password)
        self.role_combo = QComboBox()
        self.role_combo.addItem("Data Entry (No Reports Access)", "data_entry")
        self.role_combo.addItem("Administrator (Full Access)", "admin")

        form.addRow("Username / اسم المستخدم:", self.u_input)
        form.addRow("Password / كلمة المرور:", self.p_input)
        form.addRow("Role / الصلاحية:", self.role_combo)
        layout.addLayout(form)

        btn_create = QPushButton("Create New User / إنشاء مستخدم جديد")
        btn_create.setStyleSheet("background-color: #0d6efd; color: white; font-weight: bold; padding: 6px;")
        btn_create.clicked.connect(self.create_user)
        layout.addWidget(btn_create)

        layout.addWidget(QLabel("<b>Registered Users / المستخدمون المسجلون:</b>"))
        self.user_table = QTableWidget()
        self.user_table.setColumnCount(4)
        self.user_table.setHorizontalHeaderLabels(["ID", "Username", "Role", "Password Reset"])
        self.user_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.user_table)

        btn_delete = QPushButton("Delete Selected User / حذف المستخدم المحدد")
        btn_delete.setStyleSheet("background-color: #dc3545; color: white; font-weight: bold;")
        btn_delete.clicked.connect(self.delete_user)
        layout.addWidget(btn_delete)

        self.setLayout(layout)
        self.load_users()

    def load_users(self):
        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        c.execute("SELECT id, username, role FROM users")
        rows = c.fetchall()
        conn.close()

        self.user_table.setRowCount(0)
        for idx, (uid, uname, role) in enumerate(rows):
            self.user_table.insertRow(idx)
            self.user_table.setItem(idx, 0, QTableWidgetItem(str(uid)))
            self.user_table.setItem(idx, 1, QTableWidgetItem(uname))
            self.user_table.setItem(idx, 2, QTableWidgetItem("Admin" if role == 'admin' else "Data Entry"))

            btn_change_pwd = QPushButton("🔑 Change Password")
            btn_change_pwd.setStyleSheet("background-color: #ffc107; color: black; font-weight: bold; padding: 4px;")
            btn_change_pwd.clicked.connect(lambda checked, u=uname: self.open_change_password_dialog(u))
            self.user_table.setCellWidget(idx, 3, btn_change_pwd)

    def create_user(self):
        u = self.u_input.text().strip()
        p = self.p_input.text().strip()
        r = self.role_combo.currentData()
        if not u or not p:
            QMessageBox.warning(self, "Error", "Username and Password cannot be empty!")
            return
        p_hash = bcrypt.hashpw(p.encode(), bcrypt.gensalt()).decode()
        conn = sqlite3.connect("import_enterprise.db")
        try:
            conn.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)", (u, p_hash, r))
            conn.commit()
            record_log(self.current_user, "CREATE_USER", f"Created user '{u}' with role '{r}'")
            QMessageBox.information(self, "Success", f"User '{u}' created successfully.")
            self.u_input.clear()
            self.p_input.clear()
            self.load_users()
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Error", "Username already exists!")
        finally:
            conn.close()

    def open_change_password_dialog(self, target_username):
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Change Password: {target_username}")
        dialog.resize(340, 160)
        d_layout = QVBoxLayout()
        d_form = QFormLayout()

        pwd_new = QLineEdit()
        pwd_new.setEchoMode(QLineEdit.Password)
        pwd_confirm = QLineEdit()
        pwd_confirm.setEchoMode(QLineEdit.Password)

        d_form.addRow("New Password:", pwd_new)
        d_form.addRow("Confirm Password:", pwd_confirm)
        d_layout.addLayout(d_form)

        btn_save = QPushButton("Save New Password")
        btn_save.setStyleSheet("background-color: #198754; color: white; font-weight: bold; padding: 6px;")

        def apply_new_password():
            p1 = pwd_new.text().strip()
            p2 = pwd_confirm.text().strip()
            if not p1:
                QMessageBox.warning(dialog, "Error", "Password cannot be blank!")
                return
            if p1 != p2:
                QMessageBox.warning(dialog, "Error", "Passwords do not match!")
                return

            new_hash = bcrypt.hashpw(p1.encode(), bcrypt.gensalt()).decode()
            conn = sqlite3.connect("import_enterprise.db")
            conn.execute("UPDATE users SET password_hash = ? WHERE username = ?", (new_hash, target_username))
            conn.commit()
            conn.close()

            record_log(self.current_user, "CHANGE_PASSWORD", f"Changed password for user '{target_username}'")
            QMessageBox.information(dialog, "Success", f"Password for '{target_username}' updated successfully!")
            dialog.accept()

        btn_save.clicked.connect(apply_new_password)
        d_layout.addWidget(btn_save)
        dialog.setLayout(d_layout)
        dialog.exec_()

    def delete_user(self):
        row = self.user_table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Notice", "Select a user to delete.")
            return
        uid = int(self.user_table.item(row, 0).text())
        uname = self.user_table.item(row, 1).text()

        if uname == self.current_user:
            QMessageBox.warning(self, "Error", "Cannot delete your own logged-in account!")
            return

        conn = sqlite3.connect("import_enterprise.db")
        conn.execute("DELETE FROM users WHERE id = ?", (uid,))
        conn.commit()
        conn.close()
        record_log(self.current_user, "DELETE_USER", f"Deleted user '{uname}' (ID: {uid})")
        self.load_users()

# -------------------------------------------------------------
# دليل الجهات
# -------------------------------------------------------------
class EntityManagerDialog(QDialog):
    def __init__(self, current_user, refresh_callback, default_type='CLIENT'):
        super().__init__()
        self.current_user = current_user
        self.refresh_callback = refresh_callback
        self.setWindowTitle(f"{APP_NAME} - Entities Directory")
        self.resize(520, 400)
        layout = QVBoxLayout()

        form = QFormLayout()
        self.name_input = QLineEdit()
        self.phone_input = QLineEdit()
        self.type_combo = QComboBox()
        self.type_combo.addItem("Client", "CLIENT")
        self.type_combo.addItem("Supplier (Factory)", "SUPPLIER")
        self.type_combo.addItem("Shipping Agent", "SHIPPING_AGENT")

        idx = self.type_combo.findData(default_type)
        if idx != -1:
            self.type_combo.setCurrentIndex(idx)

        form.addRow("Name:", self.name_input)
        form.addRow("Phone:", self.phone_input)
        form.addRow("Type:", self.type_combo)
        layout.addLayout(form)

        btn_add = QPushButton("Save Entity")
        btn_add.setStyleSheet("background-color: #198754; color: white; font-weight: bold; padding: 6px;")
        btn_add.clicked.connect(self.save_entity)
        layout.addWidget(btn_add)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["ID", "Name", "Type", "Phone"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        btn_delete = QPushButton("Delete Selected")
        btn_delete.setStyleSheet("background-color: #dc3545; color: white;")
        btn_delete.clicked.connect(self.delete_selected)
        layout.addWidget(btn_delete)

        self.setLayout(layout)
        self.load_entities()

    def save_entity(self):
        name = self.name_input.text().strip()
        phone = self.phone_input.text().strip()
        e_type = self.type_combo.currentData()
        if not name:
            return
        conn = sqlite3.connect("import_enterprise.db")
        conn.execute("INSERT INTO entities (name, entity_type, phone) VALUES (?, ?, ?)", (name, e_type, phone))
        conn.commit()
        conn.close()
        record_log(self.current_user, "ADD_ENTITY", f"Added entity '{name}' type '{e_type}'")
        self.name_input.clear()
        self.phone_input.clear()
        self.load_entities()
        self.refresh_callback()

    def load_entities(self):
        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        c.execute("SELECT id, name, entity_type, phone FROM entities ORDER BY id DESC")
        rows = c.fetchall()
        conn.close()
        self.table.setRowCount(0)
        for r_idx, r in enumerate(rows):
            self.table.insertRow(r_idx)
            self.table.setItem(r_idx, 0, QTableWidgetItem(str(r[0])))
            self.table.setItem(r_idx, 1, QTableWidgetItem(r[1]))
            self.table.setItem(r_idx, 2, QTableWidgetItem(r[2]))
            self.table.setItem(r_idx, 3, QTableWidgetItem(r[3] or ""))

    def delete_selected(self):
        curr_row = self.table.currentRow()
        if curr_row < 0:
            return
        e_id = int(self.table.item(curr_row, 0).text())
        e_name = self.table.item(curr_row, 1).text()
        conn = sqlite3.connect("import_enterprise.db")
        conn.execute("DELETE FROM entities WHERE id = ?", (e_id,))
        conn.commit()
        conn.close()
        record_log(self.current_user, "DELETE_ENTITY", f"Deleted entity '{e_name}' (ID: {e_id})")
        self.load_entities()
        self.refresh_callback()

# -------------------------------------------------------------
# تسجيل الدخول
# -------------------------------------------------------------
class LoginDialog(QWidget):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback
        self.setWindowTitle(f"{APP_NAME} - System Access")
        self.resize(380, 240)
        layout = QVBoxLayout()

        lbl_title = QLabel(f"<b>{APP_NAME} ERP System</b>")
        lbl_title.setAlignment(Qt.AlignCenter)
        lbl_title.setStyleSheet("font-size: 16px; color: #0d6efd; margin-bottom: 5px;")
        layout.addWidget(lbl_title)

        form = QFormLayout()
        self.u = QLineEdit()
        self.p = QLineEdit()
        self.p.setEchoMode(QLineEdit.Password)
        form.addRow("Username:", self.u)
        form.addRow("Password:", self.p)
        layout.addLayout(form)

        btn = QPushButton("Login")
        btn.setStyleSheet("background-color: #0d6efd; color: white; padding: 7px; font-weight: bold;")
        btn.clicked.connect(self.auth)
        layout.addWidget(btn)

        lbl_rights = QLabel(COPYRIGHT_TEXT)
        lbl_rights.setAlignment(Qt.AlignCenter)
        lbl_rights.setStyleSheet("color: #6c757d; font-size: 10px; margin-top: 8px;")
        layout.addWidget(lbl_rights)
        self.setLayout(layout)

    def auth(self):
        u = self.u.text().strip()
        p = self.p.text().strip()
        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        c.execute("SELECT password_hash, role FROM users WHERE username = ?", (u,))
        row = c.fetchone()
        conn.close()
        if row and bcrypt.checkpw(p.encode(), row[0].encode()):
            record_log(u, "LOGIN", f"User logged in with role '{row[1]}'")
            self.close()
            self.callback(u, row[1])
        else:
            QMessageBox.warning(self, "Login Error", "Invalid Username or Password!")

# -------------------------------------------------------------
# الشاشة الرئيسية والتلميحات التوضيحية
# -------------------------------------------------------------
class MainEnterpriseApp(QMainWindow):
    def __init__(self, username, role):
        super().__init__()
        self.username = username
        self.role = role
        self.current_lang = 'en'
        self.current_year = datetime.now().year
        self._is_calculating = False
        self._is_cnt_calculating = False
        self.editing_invoice_id = None

        self.resize(1400, 940)
        self.build_ui()
        self.refresh_all_dropdowns()
        self.populate_fiscal_years()
        self.load_registered_invoices()
        self.apply_language(self.current_lang)

    def build_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # شريط الحالة
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        lbl_rights = QLabel(f"{APP_NAME} | {COPYRIGHT_TEXT}")
        lbl_rights.setStyleSheet("color: #495057; font-weight: bold; padding: 2px 8px;")
        self.status.addPermanentWidget(lbl_rights)

        # شريط التحكم العلوي
        self.nav_bar = QHBoxLayout()

        self.btn_file = QPushButton()
        self.btn_file.setStyleSheet("padding: 7px 14px; font-weight: bold; background-color: #f8f9fa; border: 1px solid #ced4da; border-radius: 4px;")
        file_menu = QMenu(self)
        self.act_entities = file_menu.addAction("")
        self.act_entities.triggered.connect(lambda: self.open_entity_manager('CLIENT'))
        if self.role == "admin":
            self.act_users = file_menu.addAction("")
            self.act_users.triggered.connect(self.open_user_manager)
            self.act_locks = file_menu.addAction("Fiscal Year Lock Management")
            self.act_locks.triggered.connect(self.open_fiscal_lock_manager)
            self.act_logs = file_menu.addAction("")
            self.act_logs.triggered.connect(lambda: AuditLogDialog(self.username).exec_())
            self.act_supplier_accounts = file_menu.addAction("Supplier Accounts")
            self.act_supplier_accounts.triggered.connect(self.open_supplier_account)
        self.act_about = file_menu.addAction("")
        self.act_about.triggered.connect(self.show_about_dialog)
        file_menu.addSeparator()
        act_exit = file_menu.addAction("Exit")
        act_exit.triggered.connect(self.close)
        self.btn_file.setMenu(file_menu)
        self.nav_bar.addWidget(self.btn_file)

        self.btn_nav_pur = QPushButton()
        self.btn_nav_ship = QPushButton()
        self.btn_nav_pay = QPushButton()
        self.nav_buttons = [self.btn_nav_pur, self.btn_nav_ship, self.btn_nav_pay]

        self.btn_nav_pur.clicked.connect(lambda: self.switch_view(0))
        self.btn_nav_ship.clicked.connect(lambda: self.switch_view(1))
        self.btn_nav_pay.clicked.connect(lambda: self.switch_view(2))

        for btn in self.nav_buttons:
            btn.setStyleSheet("padding: 7px 14px; font-weight: bold;")
            self.nav_bar.addWidget(btn)

        if self.role == "admin":
            self.btn_nav_rep = QPushButton()
            self.btn_nav_cnt_rep = QPushButton()
            self.btn_nav_client_dossier = QPushButton()

            for b in [self.btn_nav_rep, self.btn_nav_cnt_rep, self.btn_nav_client_dossier]:
                b.setStyleSheet("padding: 7px 14px; font-weight: bold;")
                self.nav_buttons.append(b)
                self.nav_bar.addWidget(b)

            self.btn_nav_rep.clicked.connect(lambda: self.switch_view(3))
            self.btn_nav_cnt_rep.clicked.connect(lambda: self.switch_view(4))
            self.btn_nav_client_dossier.clicked.connect(lambda: self.switch_view(5))

        self.nav_bar.addStretch()

        lbl_globe = QLabel("🌐")
        self.combo_lang = QComboBox()
        self.combo_lang.setStyleSheet("padding: 4px 10px; font-weight: bold;")
        self.combo_lang.addItem("English", "en")
        self.combo_lang.addItem("العربية", "ar")
        self.combo_lang.addItem("中文", "zh")
        self.combo_lang.currentIndexChanged.connect(self.on_lang_changed)
        self.nav_bar.addWidget(lbl_globe)
        self.nav_bar.addWidget(self.combo_lang)

        main_layout.addLayout(self.nav_bar)

        self.stack = QStackedWidget()
        main_layout.addWidget(self.stack)

        # ---------------- 1. شاشة المشتريات ----------------
        self.page_pur = QWidget()
        l_pur = QVBoxLayout()
        self.box_pur = QGroupBox()
        f_pur = QFormLayout()

        h1 = QHBoxLayout()
        self.lbl_inv_num = QLabel()
        self.pur_inv_num = QLineEdit()
        self.lbl_client1 = QLabel()
        self.pur_client = QComboBox()
        btn_c1 = QPushButton("+")
        btn_c1.setFixedWidth(30)
        btn_c1.clicked.connect(lambda: self.open_entity_manager('CLIENT'))

        self.lbl_supp = QLabel()
        self.pur_supplier = QComboBox()
        btn_s = QPushButton("+")
        btn_s.setFixedWidth(30)
        btn_s.clicked.connect(lambda: self.open_entity_manager('SUPPLIER'))

        h1.addWidget(self.lbl_inv_num)
        h1.addWidget(self.pur_inv_num)
        h1.addWidget(self.lbl_client1)
        h1.addWidget(self.pur_client)
        h1.addWidget(btn_c1)
        h1.addWidget(self.lbl_supp)
        h1.addWidget(self.pur_supplier)
        h1.addWidget(btn_s)
        f_pur.addRow(h1)

        h2 = QHBoxLayout()
        self.pur_curr = QComboBox()
        self.pur_curr.addItems(["USD", "RMB", "EUR", "EGP"])
        self.pur_curr.currentIndexChanged.connect(self.on_pur_currency_changed)

        self.pur_fx_rate = QLineEdit("50.0")
        self.pur_fx_rate.setPlaceholderText("FX Rate")
        self.pur_fx_rate.textChanged.connect(self.on_pur_total_or_fx_changed)

        self.lbl_inv_tot = QLabel()
        self.pur_total = QLineEdit()
        self.pur_total.setPlaceholderText("0.00")
        self.pur_total.textChanged.connect(self.on_pur_total_or_fx_changed)

        self.lbl_inv_net = QLabel()
        self.pur_net = QLineEdit()
        self.pur_net.setPlaceholderText("0.00 EGP")
        self.pur_net.setReadOnly(True)

        self.lbl_inv_cbm = QLabel()
        self.pur_cbm = QLineEdit()

        h2.addWidget(QLabel("Currency:"))
        h2.addWidget(self.pur_curr)
        h2.addWidget(QLabel("FX Rate:"))
        h2.addWidget(self.pur_fx_rate)
        h2.addWidget(self.lbl_inv_tot)
        h2.addWidget(self.pur_total)
        h2.addWidget(self.lbl_inv_net)
        h2.addWidget(self.pur_net)
        h2.addWidget(self.lbl_inv_cbm)
        h2.addWidget(self.pur_cbm)
        f_pur.addRow(h2)

        self.box_pur.setLayout(f_pur)
        l_pur.addWidget(self.box_pur)

        pur_items_bar = QHBoxLayout()
        pur_items_bar.addWidget(QLabel("<b>Invoice Cargo Items (أصناف الفاتورة):</b>"))
        
        self.btn_import_pur_excel = QPushButton("📥 Import from Excel")
        self.btn_import_pur_excel.setStyleSheet("background-color: #0dcaf0; font-weight: bold;")
        self.btn_import_pur_excel.clicked.connect(lambda: self.import_excel_to_table(self.pur_table))
        
        btn_add_pur_row = QPushButton("+ Add Item Row")
        btn_add_pur_row.setStyleSheet("background-color: #28a745; color: white; font-weight: bold;")
        btn_add_pur_row.clicked.connect(lambda: self.add_purchase_row())
        
        btn_del_pur_row = QPushButton("- Delete Selected Row")
        btn_del_pur_row.setStyleSheet("background-color: #dc3545; color: white;")
        btn_del_pur_row.clicked.connect(lambda: self.delete_table_row(self.pur_table))
        
        pur_items_bar.addStretch()
        pur_items_bar.addWidget(self.btn_import_pur_excel)
        pur_items_bar.addWidget(btn_add_pur_row)
        pur_items_bar.addWidget(btn_del_pur_row)
        l_pur.addLayout(pur_items_bar)

        self.pur_table = QTableWidget(5, 11)
        self.pur_table.setHorizontalHeaderLabels([
            "Item Code", "Item Description (AR)", "Item Description (CN)", "Cartons",
            "Pcs/Carton", "Total Pcs", "Piece Price", "Total Price",
            "CBM/Carton", "Total CBM", "Weight (KG)"
        ])
        self.pur_table.setToolTip("Maximum 15 item rows per purchase invoice.")
        for row in range(self.pur_table.rowCount()):
            for col in [5,7,9]:
                item = QTableWidgetItem("")
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.pur_table.setItem(row,col,item)
        self.pur_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.pur_table.cellChanged.connect(self.on_pur_table_cell_changed)
        l_pur.addWidget(self.pur_table)

        btn_box = QHBoxLayout()
        self.btn_save_pur = QPushButton()
        self.btn_save_pur.setStyleSheet("background-color: #0d6efd; color: white; font-weight: bold; padding: 7px;")
        self.btn_save_pur.clicked.connect(self.save_purchase_invoice)

        self.btn_cancel_pur_edit = QPushButton("Cancel Edit / إلغاء التعديل")
        self.btn_cancel_pur_edit.setStyleSheet("background-color: #6c757d; color: white; font-weight: bold; padding: 7px;")
        self.btn_cancel_pur_edit.setVisible(False)
        self.btn_cancel_pur_edit.clicked.connect(self.reset_purchase_form)

        btn_box.addWidget(self.btn_save_pur)
        btn_box.addWidget(self.btn_cancel_pur_edit)
        l_pur.addLayout(btn_box)

        l_pur.addWidget(QLabel("<b>Registered Invoices History (الفواتير المسجلة مسبقاً - يمكنك تعديل أو حذف أي فاتورة من هنا):</b>"))
        self.table_invoices_list = QTableWidget()
        self.table_invoices_list.setColumnCount(8)
        self.table_invoices_list.setHorizontalHeaderLabels([
            "ID", "Invoice No", "Client", "Supplier", "Total", "Currency", "Total Local (EGP)", "Actions"
        ])
        self.table_invoices_list.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_invoices_list.setFixedHeight(180)
        l_pur.addWidget(self.table_invoices_list)

        self.page_pur.setLayout(l_pur)
        self.stack.addWidget(self.page_pur)

        # ---------------- 2. شاشة الشحن ----------------
        self.page_ship = QWidget()
        l_ship = QVBoxLayout()

        self.box_ship = QGroupBox()
        f_ship = QFormLayout()

        r1 = QHBoxLayout()
        self.lbl_cnt_num = QLabel()
        self.ship_cnt_num = QLineEdit()
        self.ship_cnt_num.setReadOnly(True)
        self.ship_cnt_num.setPlaceholderText("Auto-generated")
        self.lbl_client2 = QLabel()
        self.ship_client = QComboBox()
        self.ship_client.currentIndexChanged.connect(self.on_shipping_client_changed)
        btn_c2 = QPushButton("+")
        btn_c2.setFixedWidth(30)
        btn_c2.clicked.connect(lambda: self.open_entity_manager('CLIENT'))

        self.lbl_agent = QLabel()
        self.ship_agent = QComboBox()
        btn_a = QPushButton("+")
        btn_a.setFixedWidth(30)
        btn_a.clicked.connect(lambda: self.open_entity_manager('SHIPPING_AGENT'))

        self.lbl_line = QLabel()
        self.ship_line = QLineEdit()

        r1.addWidget(self.lbl_cnt_num)
        r1.addWidget(self.ship_cnt_num)
        r1.addWidget(self.lbl_client2)
        r1.addWidget(self.ship_client)
        r1.addWidget(btn_c2)
        r1.addWidget(self.lbl_agent)
        r1.addWidget(self.ship_agent)
        r1.addWidget(btn_a)
        r1.addWidget(self.lbl_line)
        r1.addWidget(self.ship_line)
        f_ship.addRow(r1)

        r2 = QHBoxLayout()
        self.lbl_bol = QLabel()
        self.ship_bol = QLineEdit()

        self.ship_status = QComboBox()
        self.ship_status.addItems([
            "قيد الشحن (In-Transit)",
            "وصلت الميناء (Arrived at Port)",
            "تحت التخليص الجمركي (Under Customs)",
            "تم الإفراج والمخزنة (Cleared & Stocked)"
        ])

        r2.addWidget(self.lbl_bol)
        r2.addWidget(self.ship_bol)
        r2.addWidget(QLabel("Container Lifecycle Status:"))
        r2.addWidget(self.ship_status)
        f_ship.addRow(r2)

        r3 = QHBoxLayout()
        self.lbl_freight = QLabel()
        self.ship_freight_curr = QComboBox()
        self.ship_freight_curr.addItems(["USD $", "RMB ¥", "EUR €", "EGP ج.م"])
        self.ship_freight = QLineEdit("0.0")
        self.ship_freight.setPlaceholderText("Sea Freight Amount")

        self.ship_freight_fx = QLineEdit("50.0")
        self.ship_freight_fx.setPlaceholderText("FX Rate to EGP")

        self.ship_freight.textChanged.connect(self.recalc_container_total_egp)
        self.ship_freight_fx.textChanged.connect(self.recalc_container_total_egp)
        self.ship_freight_curr.currentIndexChanged.connect(self.on_freight_currency_changed)

        r3.addWidget(self.lbl_freight)
        r3.addWidget(self.ship_freight_curr)
        r3.addWidget(self.ship_freight)
        r3.addWidget(QLabel("Freight FX Rate (سعر الصرف):"))
        r3.addWidget(self.ship_freight_fx)
        f_ship.addRow(r3)

        r4 = QHBoxLayout()
        self.lbl_customs = QLabel()
        self.ship_customs = QLineEdit("0.0")
        self.ship_customs.setPlaceholderText("Customs in EGP")

        self.lbl_comm = QLabel()
        self.ship_comm = QLineEdit("0.0")
        self.ship_comm.setPlaceholderText("Commission in EGP")

        self.lbl_total_cnt = QLabel()
        self.ship_total_egp = QLineEdit("0.00 EGP")
        self.ship_total_egp.setReadOnly(True)
        self.ship_total_egp.setStyleSheet("font-weight: bold; background-color: #e9ecef; color: #0d6efd;")

        self.ship_customs.textChanged.connect(self.recalc_container_total_egp)
        self.ship_comm.textChanged.connect(self.recalc_container_total_egp)

        r4.addWidget(self.lbl_customs)
        r4.addWidget(self.ship_customs)
        r4.addWidget(self.lbl_comm)
        r4.addWidget(self.ship_comm)
        r4.addWidget(self.lbl_total_cnt)
        r4.addWidget(self.ship_total_egp)
        f_ship.addRow(r4)

        self.box_ship.setLayout(f_ship)
        l_ship.addWidget(self.box_ship)

        cnt_items_bar = QHBoxLayout()
        cnt_items_bar.addWidget(QLabel("<b>Container Cargo Picker (سحب البضاعة من فواتير الشراء غير المشحونة):</b>"))
        cnt_items_bar.addStretch()

        self.btn_pull_ship_cargo = QPushButton("📥 Pull Unshipped Invoices")
        self.btn_pull_ship_cargo.setStyleSheet("background-color:#0dcaf0;font-weight:bold;padding:6px;")
        self.btn_pull_ship_cargo.clicked.connect(self.open_cargo_picker)
        cnt_items_bar.addWidget(self.btn_pull_ship_cargo)
        l_ship.addLayout(cnt_items_bar)

        self.ship_items_table = QTableWidget(0, 10)
        self.ship_items_table.setHorizontalHeaderLabels([
            "Source Invoice", "Item Code", "Description (AR)", "Description (CN)",
            "Cartons", "Pcs/Carton", "Total Pcs", "Piece Price", "Total CBM", "Weight (KG)"
        ])
        self.ship_items_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.ship_items_table.setEditTriggers(QTableWidget.NoEditTriggers)
        l_ship.addWidget(self.ship_items_table)

        # Live CBM counter for all cargo currently added to this container.
        self.lbl_ship_total_cbm = QLabel("Total Container CBM: 0.00 CBM")
        self.lbl_ship_total_cbm.setStyleSheet(
            "font-size: 15px; font-weight: bold; padding: 8px;"
        )
        l_ship.addWidget(self.lbl_ship_total_cbm)

        self.btn_save_ship = QPushButton()
        self.btn_save_ship.setStyleSheet("background-color: #0d6efd; color: white; font-weight: bold; padding: 7px;")
        self.btn_save_ship.clicked.connect(self.save_container_data)
        l_ship.addWidget(self.btn_save_ship)

        self.page_ship.setLayout(l_ship)
        self.stack.addWidget(self.page_ship)

        # ---------------- 3. التسديدات والمصروفات ----------------
        self.page_pay = QWidget()
        l_pay = QVBoxLayout()
        self.sub_tabs = QTabWidget()

        self.t_client_pay = QWidget()
        l_cp = QFormLayout()
        h_cp = QHBoxLayout()
        self.pay_client = QComboBox()
        btn_c3 = QPushButton("+")
        btn_c3.setFixedWidth(30)
        btn_c3.clicked.connect(lambda: self.open_entity_manager('CLIENT'))
        h_cp.addWidget(self.pay_client)
        h_cp.addWidget(btn_c3)

        self.lbl_pay_client = QLabel()
        self.lbl_pay_amt = QLabel()
        
        h_amt = QHBoxLayout()
        self.pay_client_amt = QLineEdit()
        self.pay_curr = QComboBox()
        self.pay_curr.addItems(["EGP", "USD", "RMB", "EUR"])
        self.pay_fx = QLineEdit("1.0")
        h_amt.addWidget(self.pay_client_amt)
        h_amt.addWidget(QLabel("Currency:"))
        h_amt.addWidget(self.pay_curr)
        h_amt.addWidget(QLabel("FX Rate:"))
        h_amt.addWidget(self.pay_fx)

        self.lbl_pay_meth = QLabel()
        self.pay_client_method = QComboBox()
        self.pay_client_method.addItems(["Cash", "Bank Transfer (SWIFT)", "Credit Card"])
        self.pay_client_doc = QLineEdit()
        self.pay_client_doc.setReadOnly(True)

        self.btn_attach = QPushButton()
        self.btn_attach.clicked.connect(lambda: self.pick_file(self.pay_client_doc))
        self.btn_save_pay = QPushButton()
        self.btn_save_pay.setStyleSheet("background-color: #198754; color: white; font-weight: bold;")
        self.btn_save_pay.clicked.connect(self.save_client_payment)

        l_cp.addRow(self.lbl_pay_client, h_cp)
        l_cp.addRow(self.lbl_pay_amt, h_amt)
        l_cp.addRow(self.lbl_pay_meth, self.pay_client_method)
        l_cp.addRow(self.btn_attach, self.pay_client_doc)
        l_cp.addRow(self.btn_save_pay)
        self.t_client_pay.setLayout(l_cp)
        self.sub_tabs.addTab(self.t_client_pay, "")

        self.t_exp = QWidget()
        l_e = QFormLayout()
        self.lbl_exp_cnt = QLabel()
        self.exp_container = QComboBox()
        self.lbl_exp_cat = QLabel()
        self.exp_type = QComboBox()
        self.exp_type.addItems(["Customs & Duties", "Ocean Freight & Transport", "Port Storage & Demurrage", "Handling & Operations"])
        
        h_exp_amt = QHBoxLayout()
        self.exp_amt = QLineEdit()
        self.exp_curr = QComboBox()
        self.exp_curr.addItems(["EGP", "USD", "RMB", "EUR"])
        self.exp_fx = QLineEdit("1.0")
        h_exp_amt.addWidget(self.exp_amt)
        h_exp_amt.addWidget(QLabel("Currency:"))
        h_exp_amt.addWidget(self.exp_curr)
        h_exp_amt.addWidget(QLabel("FX Rate:"))
        h_exp_amt.addWidget(self.exp_fx)

        self.lbl_exp_notes = QLabel()
        self.exp_notes = QLineEdit()
        self.btn_save_exp = QPushButton()
        self.btn_save_exp.setStyleSheet("background-color: #dc3545; color: white; font-weight: bold;")
        self.btn_save_exp.clicked.connect(self.save_expense_data)

        l_e.addRow(self.lbl_exp_cnt, self.exp_container)
        l_e.addRow(self.lbl_exp_cat, self.exp_type)
        l_e.addRow(QLabel("Expense Amount:"), h_exp_amt)
        l_e.addRow(self.lbl_exp_notes, self.exp_notes)
        l_e.addRow(self.btn_save_exp)
        self.t_exp.setLayout(l_e)
        self.sub_tabs.addTab(self.t_exp, "")

        # ---------------- Supplier Payments ----------------
        self.t_supplier_pay = QWidget()
        l_sp = QFormLayout()
        self.pay_supplier = QComboBox()
        btn_sp = QPushButton("+")
        btn_sp.setFixedWidth(30)
        btn_sp.clicked.connect(lambda: self.open_entity_manager('SUPPLIER'))
        h_sp = QHBoxLayout()
        h_sp.addWidget(self.pay_supplier)
        h_sp.addWidget(btn_sp)

        self.supplier_pay_amt = QLineEdit()
        self.supplier_pay_curr = QComboBox()
        self.supplier_pay_curr.addItems(["EGP", "USD", "RMB", "EUR"])
        self.supplier_pay_fx = QLineEdit("1.0")
        h_sp_amt = QHBoxLayout()
        h_sp_amt.addWidget(self.supplier_pay_amt)
        h_sp_amt.addWidget(QLabel("Currency:"))
        h_sp_amt.addWidget(self.supplier_pay_curr)
        h_sp_amt.addWidget(QLabel("FX Rate:"))
        h_sp_amt.addWidget(self.supplier_pay_fx)

        self.supplier_pay_method = QComboBox()
        self.supplier_pay_method.addItems(["Cash", "Bank Transfer (SWIFT)", "Bank Transfer", "Credit Card"])
        self.supplier_pay_notes = QLineEdit()
        self.btn_save_supplier_pay = QPushButton("Save Supplier Payment")
        self.btn_save_supplier_pay.setStyleSheet("background-color:#6f42c1;color:white;font-weight:bold;")
        self.btn_save_supplier_pay.clicked.connect(self.save_supplier_payment)

        l_sp.addRow(QLabel("Supplier:"), h_sp)
        l_sp.addRow(QLabel("Payment Amount:"), h_sp_amt)
        l_sp.addRow(QLabel("Payment Method:"), self.supplier_pay_method)
        l_sp.addRow(QLabel("Notes:"), self.supplier_pay_notes)
        l_sp.addRow(self.btn_save_supplier_pay)
        self.t_supplier_pay.setLayout(l_sp)
        self.sub_tabs.addTab(self.t_supplier_pay, "Supplier Payments")

        l_pay.addWidget(self.sub_tabs)
        self.page_pay.setLayout(l_pay)
        self.stack.addWidget(self.page_pay)

        # ---------------- 4. شاشة الملخص المالي ----------------
        if self.role == "admin":
            self.page_rep = QWidget()
            l_rep = QVBoxLayout()

            filter_box = QGroupBox()
            f_layout = QHBoxLayout()
            self.lbl_fyear = QLabel()
            self.combo_fiscal_year = QComboBox()
            self.combo_fiscal_year.setStyleSheet("font-weight: bold; padding: 4px 8px;")

            self.btn_filter_year = QPushButton()
            self.btn_filter_year.setStyleSheet("background-color: #0d6efd; color: white; font-weight: bold; padding: 6px 14px;")
            self.btn_filter_year.clicked.connect(self.load_admin_reports)

            self.btn_export = QPushButton()
            self.btn_export.setStyleSheet("background-color: #198754; color: white; font-weight: bold; padding: 6px 14px;")
            self.btn_export.clicked.connect(self.export_excel_english)

            f_layout.addWidget(self.lbl_fyear)
            f_layout.addWidget(self.combo_fiscal_year)
            f_layout.addWidget(self.btn_filter_year)
            f_layout.addStretch()
            f_layout.addWidget(self.btn_export)
            filter_box.setLayout(f_layout)
            l_rep.addWidget(filter_box)

            self.rep_table = QTableWidget()
            self.rep_table.setColumnCount(5)
            self.rep_table.setHorizontalHeaderLabels([
                "Client Name", "Total Expenses (Local)", "Total Paid (Local)", "Remaining Balance", "Account Status"
            ])
            self.rep_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            l_rep.addWidget(self.rep_table)

            self.page_rep.setLayout(l_rep)
            self.stack.addWidget(self.page_rep)

            # ---------------- 5. شاشة تقرير الحاويات ----------------
            self.page_cnt_rep = QWidget()
            l_cr = QVBoxLayout()

            sel_cnt_box = QHBoxLayout()
            sel_cnt_box.addWidget(QLabel("<b>Select Container / اختر الحاوية:</b>"))
            self.combo_rep_cnt = QComboBox()
            self.combo_rep_cnt.setStyleSheet("font-weight: bold; padding: 4px 10px;")
            self.combo_rep_cnt.currentIndexChanged.connect(self.load_container_manifest)
            sel_cnt_box.addWidget(self.combo_rep_cnt)

            self.btn_update_status = QPushButton("Change Status / الترحيل للمخزن")
            self.btn_update_status.setStyleSheet("background-color: #0d6efd; color: white; font-weight: bold; padding: 6px;")
            self.btn_update_status.clicked.connect(self.mark_container_cleared)
            sel_cnt_box.addWidget(self.btn_update_status)

            self.btn_export_cnt = QPushButton()
            self.btn_export_cnt.setStyleSheet("background-color: #198754; color: white; font-weight: bold; padding: 6px 14px;")
            self.btn_export_cnt.clicked.connect(self.export_container_manifest_excel)
            sel_cnt_box.addWidget(self.btn_export_cnt)

            sel_cnt_box.addStretch()
            l_cr.addLayout(sel_cnt_box)

            self.lbl_cnt_header = QLabel("Container Info: -")
            self.lbl_cnt_header.setStyleSheet("font-size: 13px; font-weight: bold; background-color: #e9ecef; padding: 8px; border-radius: 4px;")
            l_cr.addWidget(self.lbl_cnt_header)

            l_cr.addWidget(QLabel("<b>Container Items & Landed Costing:</b>"))
            self.table_cnt_items = QTableWidget()
            self.table_cnt_items.setColumnCount(9)
            self.table_cnt_items.setHorizontalHeaderLabels([
                "Item Code", "Name (AR)", "Name (CN)", "Cartons", "Purchase Price", "CBM", "Weight", "Allocated Expenses", "Landed Cost / Unit"
            ])
            self.table_cnt_items.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            l_cr.addWidget(self.table_cnt_items)

            self.page_cnt_rep.setLayout(l_cr)
            self.stack.addWidget(self.page_cnt_rep)

            # ---------------- 6. شاشة الملف الشامل للعميل ----------------
            self.page_client_dossier = QWidget()
            l_cd = QVBoxLayout()

            top_client_sel = QHBoxLayout()
            top_client_sel.addWidget(QLabel("<b>Select Client / اختر العميل:</b>"))
            self.combo_dossier_client = QComboBox()
            self.combo_dossier_client.setStyleSheet("font-weight: bold; padding: 4px 10px;")
            self.combo_dossier_client.currentIndexChanged.connect(self.load_client_dossier)
            top_client_sel.addWidget(self.combo_dossier_client)

            self.btn_export_client = QPushButton()
            self.btn_export_client.setStyleSheet("background-color: #198754; color: white; font-weight: bold; padding: 6px 14px;")
            self.btn_export_client.clicked.connect(self.export_client_dossier_excel)
            top_client_sel.addWidget(self.btn_export_client)

            top_client_sel.addStretch()
            l_cd.addLayout(top_client_sel)

            self.lbl_client_stats = QLabel("Client Stats: Total Expenses: 0.00 | Total Paid: 0.00 | Balance: 0.00")
            self.lbl_client_stats.setStyleSheet("font-size: 13px; font-weight: bold; background-color: #cfe2ff; color: #084298; padding: 10px; border-radius: 4px;")
            l_cd.addWidget(self.lbl_client_stats)

            self.tab_client_sections = QTabWidget()

            t_cnts = QWidget()
            l_cnts = QVBoxLayout()
            self.table_client_cnts = QTableWidget()
            self.table_client_cnts.setColumnCount(6)
            self.table_client_cnts.setHorizontalHeaderLabels([
                "Container No", "B/L No", "Shipping Line", "Arrival Date", "Status", "Total Cost (Local)"
            ])
            self.table_client_cnts.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            l_cnts.addWidget(self.table_client_cnts)
            t_cnts.setLayout(l_cnts)
            self.tab_client_sections.addTab(t_cnts, "Containers Shipped")

            t_pays = QWidget()
            l_pays = QVBoxLayout()
            self.table_client_pays = QTableWidget()
            self.table_client_pays.setColumnCount(5)
            self.table_client_pays.setHorizontalHeaderLabels([
                "Date", "Amount (Local)", "Method", "Notes / Description", "Receipt Document"
            ])
            self.table_client_pays.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            l_pays.addWidget(self.table_client_pays)
            t_pays.setLayout(l_pays)
            self.tab_client_sections.addTab(t_pays, "Payment History")

            t_itms = QWidget()
            l_itms = QVBoxLayout()
            self.table_client_all_items = QTableWidget()
            self.table_client_all_items.setColumnCount(8)
            self.table_client_all_items.setHorizontalHeaderLabels([
                "Container No", "Item Code", "Name (AR)", "Cartons", "CBM", "Weight", "Allocated Exp", "Landed Cost / Unit"
            ])
            self.table_client_all_items.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            l_itms.addWidget(self.table_client_all_items)
            t_itms.setLayout(l_itms)
            self.tab_client_sections.addTab(t_itms, "Imported Cargo & Landed Costs")

            l_cd.addWidget(self.tab_client_sections)
            self.page_client_dossier.setLayout(l_cd)
            self.stack.addWidget(self.page_client_dossier)

        self.switch_view(0)

    # ---------------- دوال الحساب اللحظي وتحديث حالة الحاوية ----------------
    def on_freight_currency_changed(self):
        curr_text = self.ship_freight_curr.currentText()
        if "USD" in curr_text:
            self.ship_freight_fx.setText("50.0")
        elif "RMB" in curr_text:
            self.ship_freight_fx.setText("7.0")
        elif "EUR" in curr_text:
            self.ship_freight_fx.setText("54.0")
        elif "EGP" in curr_text:
            self.ship_freight_fx.setText("1.0")
        self.recalc_container_total_egp()

    def recalc_container_total_egp(self):
        if self._is_cnt_calculating:
            return
        try:
            freight_str = self.ship_freight.text().replace(',', '').strip()
            fx_str = self.ship_freight_fx.text().replace(',', '').strip()
            customs_str = self.ship_customs.text().replace(',', '').strip()
            comm_str = self.ship_comm.text().replace(',', '').strip()

            freight = float(freight_str) if freight_str else 0.0
            fx = float(fx_str) if fx_str else 1.0
            customs = float(customs_str) if customs_str else 0.0
            comm = float(comm_str) if comm_str else 0.0

            total_egp = (freight * fx) + customs + comm
            self._is_cnt_calculating = True
            self.ship_total_egp.setText(f"{total_egp:,.2f} EGP")
            self._is_cnt_calculating = False
        except ValueError:
            pass

    # ---------------- إدارة وتعديل فواتير المشتريات ----------------
    def load_registered_invoices(self):
        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        c.execute("""
        SELECT i.id, i.invoice_num, c.name, s.name, i.total_amount, i.currency, i.total_amount_local
        FROM client_invoices i
        LEFT JOIN entities c ON i.client_id = c.id
        LEFT JOIN entities s ON i.supplier_id = s.id
        ORDER BY i.id DESC
        """)
        rows = c.fetchall()
        conn.close()

        self.table_invoices_list.setRowCount(0)
        for idx, (inv_id, num, cl_name, sup_name, tot, curr, tot_loc) in enumerate(rows):
            self.table_invoices_list.insertRow(idx)
            self.table_invoices_list.setItem(idx, 0, QTableWidgetItem(str(inv_id)))
            self.table_invoices_list.setItem(idx, 1, QTableWidgetItem(num))
            self.table_invoices_list.setItem(idx, 2, QTableWidgetItem(cl_name or ""))
            self.table_invoices_list.setItem(idx, 3, QTableWidgetItem(sup_name or ""))
            self.table_invoices_list.setItem(idx, 4, QTableWidgetItem(f"{tot:,.2f}"))
            self.table_invoices_list.setItem(idx, 5, QTableWidgetItem(curr or "USD"))
            self.table_invoices_list.setItem(idx, 6, QTableWidgetItem(f"{tot_loc:,.2f}"))

            act_widget = QWidget()
            h_act = QHBoxLayout(act_widget)
            h_act.setContentsMargins(2, 2, 2, 2)

            btn_edit = QPushButton("✏️ Edit")
            btn_edit.setStyleSheet("background-color: #ffc107; color: black; font-weight: bold; padding: 3px 6px;")
            btn_edit.clicked.connect(lambda checked, iid=inv_id: self.start_edit_invoice(iid))

            btn_del = QPushButton("🗑️")
            btn_del.setStyleSheet("background-color: #dc3545; color: white; padding: 3px 6px;")
            btn_del.clicked.connect(lambda checked, iid=inv_id, n=num: self.delete_invoice(iid, n))

            h_act.addWidget(btn_edit)
            h_act.addWidget(btn_del)
            self.table_invoices_list.setCellWidget(idx, 7, act_widget)

    def start_edit_invoice(self, invoice_id):
        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        c.execute("""SELECT id, invoice_num, client_id, supplier_id, currency, exchange_rate, total_amount, net_amount, total_cbm, fiscal_year
                     FROM client_invoices WHERE id = ?""", (invoice_id,))
        inv = c.fetchone()
        if not inv:
            conn.close()
            return

        if is_year_locked(inv[9]):
            conn.close()
            QMessageBox.critical(self, "Fiscal Lock", f"Fiscal Year {inv[9]} is LOCKED! Cannot modify this invoice.")
            return

        c.execute("SELECT COUNT(*) FROM invoice_items WHERE invoice_id=? AND COALESCE(shipped_status,0)=1", (invoice_id,))
        if c.fetchone()[0] > 0:
            conn.close()
            QMessageBox.warning(
                self, "Invoice Locked",
                "This purchase invoice contains cargo that has already been shipped. "
                "Create a new invoice for additional changes instead of editing the shipped invoice."
            )
            return

        self.editing_invoice_id = invoice_id
        self.pur_inv_num.setText(inv[1])
        idx_c = self.pur_client.findData(inv[2])
        if idx_c != -1: self.pur_client.setCurrentIndex(idx_c)
        idx_s = self.pur_supplier.findData(inv[3])
        if idx_s != -1: self.pur_supplier.setCurrentIndex(idx_s)
        idx_curr = self.pur_curr.findText(inv[4] or "USD")
        if idx_curr != -1: self.pur_curr.setCurrentIndex(idx_curr)
        self.pur_fx_rate.setText(str(inv[5]))

        c.execute("""SELECT item_code,item_ar,item_cn,cartons,pcs_per_carton,total_pcs,
                            piece_price,total_price,cbm_per_carton,total_cbm,weight
                     FROM invoice_items WHERE invoice_id=? ORDER BY id""",(invoice_id,))
        items=c.fetchall()
        conn.close()

        self.pur_table.setRowCount(0)
        for r_idx,itm in enumerate(items[:15]):
            self.pur_table.insertRow(r_idx)
            for c_idx,val in enumerate(itm):
                item=QTableWidgetItem(str(val if val is not None else ""))
                if c_idx in [5,7,9]:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.pur_table.setItem(r_idx,c_idx,item)
        while self.pur_table.rowCount() < 5:
            self.pur_table.insertRow(self.pur_table.rowCount())
        self.recalc_purchase_totals()

        t=TRANSLATIONS[self.current_lang]
        self.btn_save_pur.setText(t.get('update_inv_btn','Update Purchase Invoice'))
        self.btn_save_pur.setStyleSheet("background-color:#ffc107;color:black;font-weight:bold;padding:7px;")
    def reset_purchase_form(self):
        self.editing_invoice_id=None
        self.pur_inv_num.clear()
        self.pur_table.clearContents()
        self.pur_table.setRowCount(5)
        for r in range(self.pur_table.rowCount()):
            for col in [5,7,9]:
                item=QTableWidgetItem("")
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.pur_table.setItem(r,col,item)
        self.pur_total.clear()
        self.pur_net.clear()
        self.pur_cbm.clear()
        t=TRANSLATIONS[self.current_lang]
        self.btn_save_pur.setText(t['save_inv_btn'])
        self.btn_save_pur.setStyleSheet("background-color:#0d6efd;color:white;font-weight:bold;padding:7px;")
    def delete_invoice(self, invoice_id, invoice_num):
        conn=sqlite3.connect("import_enterprise.db")
        c=conn.cursor()
        c.execute("SELECT fiscal_year FROM client_invoices WHERE id=?",(invoice_id,))
        row=c.fetchone()
        if row and is_year_locked(row[0]):
            conn.close()
            QMessageBox.critical(self,"Fiscal Lock",f"Fiscal Year {row[0]} is LOCKED! Cannot delete invoice.")
            return
        reply=QMessageBox.question(self,"Confirm Delete",f"Are you sure you want to delete invoice '{invoice_num}'?",QMessageBox.Yes|QMessageBox.No)
        if reply != QMessageBox.Yes:
            conn.close(); return
        c.execute("SELECT COUNT(*) FROM invoice_items WHERE invoice_id=? AND COALESCE(shipped_status,0)=1",(invoice_id,))
        if c.fetchone()[0] > 0:
            conn.close()
            QMessageBox.warning(self,"Cannot Delete","This invoice has already-shipped cargo and cannot be deleted.")
            return
        c.execute("DELETE FROM supplier_ledger WHERE invoice_id=?",(invoice_id,))
        c.execute("DELETE FROM invoice_items WHERE invoice_id=?",(invoice_id,))
        c.execute("DELETE FROM client_invoices WHERE id=?",(invoice_id,))
        conn.commit(); conn.close()
        record_log(self.username,"DELETE_INVOICE",f"Deleted invoice '{invoice_num}' (ID: {invoice_id})")
        QMessageBox.information(self,"Deleted",f"Invoice '{invoice_num}' deleted successfully.")
        if self.editing_invoice_id==invoice_id: self.reset_purchase_form()
    def save_purchase_invoice(self):
        if is_year_locked(self.current_year):
            QMessageBox.critical(self,"Fiscal Lock",f"Fiscal Year {self.current_year} is LOCKED! Cannot insert or modify records.")
            return
        num=self.pur_inv_num.text().strip()
        c_id=self.pur_client.currentData()
        s_id=self.pur_supplier.currentData()
        curr=self.pur_curr.currentText()
        try:
            fx=float(self.pur_fx_rate.text().replace(",","").strip() or 1)
        except ValueError:
            QMessageBox.warning(self,"Warning","Invalid exchange rate."); return
        if not num or not c_id or not s_id:
            QMessageBox.warning(self,"Warning","Invoice Number, Client and Supplier are required.")
            return

        rows=[]
        total=0.0
        total_cbm=0.0
        for row in range(self.pur_table.rowCount()):
            code=self.pur_table.item(row,0)
            ar=self.pur_table.item(row,1)
            cn=self.pur_table.item(row,2)
            if not ar or not ar.text().strip():
                continue
            def numv(col, default=0):
                try:
                    return float(self.pur_table.item(row,col).text().replace(",","").strip() or default)
                except Exception:
                    return float(default)
            cartons=int(numv(3))
            ppc=max(1,int(numv(4,1)))
            total_pcs=int(numv(5,cartons*ppc))
            price=numv(6)
            total_price=total_pcs*price
            cbm_carton=numv(8)
            total_cbm=cartons*cbm_carton
            weight=numv(10)
            rows.append((code.text().strip() if code else "",ar.text().strip(),cn.text().strip() if cn else "",
                         cartons,ppc,total_pcs,price,total_price,cbm_carton,total_cbm,weight))
            total += total_price
            total_cbm_sum = locals().get("total_cbm_sum",0.0) + total_cbm
            locals()["total_cbm_sum"]=total_cbm_sum

        total_cbm=locals().get("total_cbm_sum",0.0)
        if not rows:
            QMessageBox.warning(self,"Warning","Add at least one purchase item.")
            return

        if len(rows)>15:
            QMessageBox.warning(self,"Maximum Items","Maximum 15 items per purchase invoice. Save this invoice and start a new invoice.")
            return

        tot_local=total*fx
        net=tot_local

        conn=sqlite3.connect("import_enterprise.db")
        try:
            c=conn.cursor()
            # Duplicate supplier invoice number is checked per supplier.
            c.execute("""SELECT id FROM client_invoices
                         WHERE supplier_id=? AND invoice_num=? AND id<>COALESCE(?,0)""",
                      (s_id,num,self.editing_invoice_id))
            if c.fetchone():
                raise ValueError("DUPLICATE_SUPPLIER_INVOICE")

            if self.editing_invoice_id:
                target=self.editing_invoice_id
                c.execute("""UPDATE client_invoices SET invoice_num=?,client_id=?,supplier_id=?,
                             currency=?,exchange_rate=?,total_amount=?,total_amount_local=?,net_amount=?,total_cbm=?
                             WHERE id=?""",(num,c_id,s_id,curr,fx,total,tot_local,net,total_cbm,target))
                c.execute("DELETE FROM invoice_items WHERE invoice_id=? AND COALESCE(shipped_status,0)=0",(target,))
                c.execute("DELETE FROM supplier_ledger WHERE invoice_id=? AND tx_type='INVOICE'",(target,))
                msg="Invoice updated successfully."
            else:
                c.execute("""INSERT INTO client_invoices
                             (invoice_num,client_id,supplier_id,currency,exchange_rate,total_amount,total_amount_local,net_amount,total_cbm,fiscal_year)
                             VALUES (?,?,?,?,?,?,?,?,?,?)""",
                          (num,c_id,s_id,curr,fx,total,tot_local,net,total_cbm,self.current_year))
                target=c.lastrowid
                msg="Invoice registered successfully."

            for row in rows:
                c.execute("""INSERT INTO invoice_items
                    (invoice_id,item_code,item_ar,item_cn,cartons,pcs_per_carton,total_pcs,piece_price,total_price,
                     cbm_per_carton,total_cbm,cbm,weight,shipped_status)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,0)""",(target,*row[:9],row[9],row[9],row[10]))

            c.execute("""INSERT INTO supplier_ledger
                (supplier_id,invoice_id,tx_type,amount,amount_local,currency,exchange_rate,supplier_invoice_num,reason,tx_date,method,created_by)
                VALUES (?,?, 'INVOICE',?,?,?,?,?,?,?, ?,?)""",
                (s_id,target,total,tot_local,curr,fx,num,f"Supplier Invoice #{num}",datetime.now().strftime("%Y-%m-%d"),"INVOICE",self.username))

            conn.commit()
            record_log(self.username,"SAVE_INVOICE",f"Saved supplier invoice #{num}: {total:,.2f} {curr}")
            QMessageBox.information(self,"Success",msg)
            self.reset_purchase_form()
            self.load_registered_invoices()
        except ValueError as e:
            conn.rollback()
            if str(e)=="DUPLICATE_SUPPLIER_INVOICE":
                QMessageBox.warning(self,"Duplicate Supplier Invoice",f"Supplier invoice number '{num}' already exists for this supplier.")
            else:
                QMessageBox.warning(self,"Validation Error",str(e))
        except sqlite3.IntegrityError as e:
            conn.rollback()
            QMessageBox.warning(self,"Error",f"Could not save invoice: {e}")
    def on_pur_currency_changed(self):
        curr = self.pur_curr.currentText()
        if curr == "USD":
            self.pur_fx_rate.setText("50.0")
        elif curr == "RMB":
            self.pur_fx_rate.setText("7.0")
        elif curr == "EUR":
            self.pur_fx_rate.setText("54.0")
        elif curr == "EGP":
            self.pur_fx_rate.setText("1.0")
        self.on_pur_total_or_fx_changed()

    def on_pur_total_or_fx_changed(self):
        if self._is_calculating:
            return
        try:
            tot_str = self.pur_total.text().replace(',', '').strip()
            fx_str = self.pur_fx_rate.text().replace(',', '').strip()
            
            tot = float(tot_str) if tot_str else 0.0
            fx = float(fx_str) if fx_str else 1.0
            
            net_val = tot * fx
            self._is_calculating = True
            self.pur_net.setText(f"{net_val:,.2f}")
            self._is_calculating = False
        except ValueError:
            pass

    def on_pur_table_cell_changed(self,row,column):
        if self._is_calculating:
            return
        if column in [3,4,6,8]:
            self.recalc_purchase_totals()

    def recalc_purchase_totals(self):
        if self._is_calculating: return
        self._is_calculating=True
        total=0.0
        total_cbm=0.0
        try:
            fx=float(self.pur_fx_rate.text().replace(",","").strip() or 1)
        except Exception:
            fx=1.0
        try:
            for r in range(self.pur_table.rowCount()):
                cartons=float((self.pur_table.item(r,3).text() if self.pur_table.item(r,3) else "0").replace(",","") or 0)
                ppc=max(1,int(float((self.pur_table.item(r,4).text() if self.pur_table.item(r,4) else "1") or 1)))
                price=float((self.pur_table.item(r,6).text() if self.pur_table.item(r,6) else "0").replace(",","") or 0)
                cbm_carton=float((self.pur_table.item(r,8).text() if self.pur_table.item(r,8) else "0").replace(",","") or 0)
                total_pcs=int(cartons*ppc)
                row_total=total_pcs*price
                row_cbm=cartons*cbm_carton
                self._set_calc_cell(r,5,total_pcs)
                self._set_calc_cell(r,7,row_total)
                self._set_calc_cell(r,9,row_cbm)
                total+=row_total
                total_cbm+=row_cbm
            self.pur_total.setText(f"{total:,.2f}")
            self.pur_net.setText(f"{total*fx:,.2f}")
            self.pur_cbm.setText(f"{total_cbm:,.2f}")
        finally:
            self._is_calculating=False

    def _set_calc_cell(self,row,col,value):
        item=self.pur_table.item(row,col)
        if item is None:
            item=QTableWidgetItem("")
            self.pur_table.setItem(row,col,item)
        item.setText(f"{value:,.2f}" if col in [7,9] else f"{int(value)}")
    def import_excel_to_table(self, table):
        path,_=QFileDialog.getOpenFileName(self,"Select Packing List Excel","","Excel Files (*.xlsx *.xls)")
        if not path: return
        try:
            wb=openpyxl.load_workbook(path,data_only=True)
            ws=wb.active
            table.setRowCount(0)
            row_count=0
            for r_idx,row in enumerate(ws.iter_rows(values_only=True)):
                if r_idx==0 or not any(row): continue
                if table==self.pur_table and row_count>=15:
                    break
                table.insertRow(row_count)
                if table==self.pur_table:
                    # New format: code, ar, cn, cartons, pcs/carton, total pcs, price, total price, cbm/carton, total cbm, weight
                    vals=list(row)
                    mapping=[0,1,2,3,4,None,6,None,8,None,10]
                    for c_idx,src in enumerate(mapping):
                        if src is not None and src<len(vals):
                            self.pur_table.setItem(row_count,c_idx,QTableWidgetItem(str(vals[src]) if vals[src] is not None else ""))
                    # Backward-compatible old 7-column import: code, ar, cn, cartons, price, total cbm, weight
                    if len(vals)<=7:
                        self.pur_table.setItem(row_count,4,QTableWidgetItem("1"))
                        self.pur_table.setItem(row_count,6,QTableWidgetItem(str(vals[4] or "")))
                        cartons=float(vals[3] or 0)
                        old_cbm=float(vals[5] or 0)
                        cbm_per_carton=(old_cbm/cartons) if cartons else 0
                        self.pur_table.setItem(row_count,8,QTableWidgetItem(str(cbm_per_carton)))
                    for col in [5,7,9]:
                        item=self.pur_table.item(row_count,col) or QTableWidgetItem("")
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        self.pur_table.setItem(row_count,col,item)
                else:
                    for c_idx in range(min(len(row),table.columnCount())):
                        table.setItem(row_count,c_idx,QTableWidgetItem(str(row[c_idx]) if row[c_idx] is not None else ""))
                row_count+=1
            if table==self.pur_table: self.recalc_purchase_totals()
            record_log(self.username,"IMPORT_EXCEL",f"Imported {row_count} rows from '{os.path.basename(path)}'")
            QMessageBox.information(self,"Success",f"Successfully imported {row_count} items from Excel.")
        except Exception as e:
            QMessageBox.warning(self,"Import Error",f"Failed to import Excel: {str(e)}")


        path, _ = QFileDialog.getOpenFileName(self, "Select Packing List Excel", "", "Excel Files (*.xlsx *.xls)")
        if not path:
            return

        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active

            table.setRowCount(0)
            row_count = 0
            for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
                if r_idx == 0:
                    continue
                if not any(row):
                    continue

                table.insertRow(row_count)
                for c_idx in range(min(len(row), 7)):
                    val = row[c_idx]
                    table.setItem(row_count, c_idx, QTableWidgetItem(str(val) if val is not None else ""))
                row_count += 1

            record_log(self.username, "IMPORT_EXCEL", f"Imported {row_count} rows from '{os.path.basename(path)}'")
            QMessageBox.information(self, "Success", f"Successfully imported {row_count} items from Excel.")
            if table == self.pur_table:
                self.recalc_purchase_totals()
        except Exception as e:
            QMessageBox.warning(self, "Import Error", f"Failed to import Excel: {str(e)}")

    def delete_table_row(self, table):
        curr=table.currentRow()
        if curr>=0:
            table.removeRow(curr)
            if table==self.pur_table:
                self.recalc_purchase_totals()
        else:
            QMessageBox.information(self,"Note","Please select a row to delete.")

    def open_fiscal_lock_manager(self):
        dlg = FiscalLockDialog(self.username)
        dlg.exec_()

    def populate_fiscal_years(self):
        if self.role != "admin" or not hasattr(self, 'combo_fiscal_year'):
            return
        self.combo_fiscal_year.clear()
        t = TRANSLATIONS[self.current_lang]
        self.combo_fiscal_year.addItem(t['all_years'], "ALL")

        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        c.execute("SELECT DISTINCT fiscal_year FROM ledger WHERE fiscal_year IS NOT NULL ORDER BY fiscal_year DESC")
        years = [row[0] for row in c.fetchall()]
        conn.close()

        if self.current_year not in years:
            years.insert(0, self.current_year)

        for y in sorted(list(set(years)), reverse=True):
            self.combo_fiscal_year.addItem(f"FY {y}", y)

        idx = self.combo_fiscal_year.findData(self.current_year)
        if idx != -1:
            self.combo_fiscal_year.setCurrentIndex(idx)

    def switch_view(self, index):
        self.stack.setCurrentIndex(index)
        for i, btn in enumerate(self.nav_buttons):
            if i == index:
                btn.setStyleSheet("padding: 7px 14px; font-weight: bold; background-color: #0d6efd; color: white; border-radius: 4px;")
            else:
                btn.setStyleSheet("padding: 7px 14px; font-weight: bold; background-color: #e9ecef; color: #333; border-radius: 4px;")

    def on_lang_changed(self):
        lang_code = self.combo_lang.currentData()
        self.apply_language(lang_code)

    def apply_language(self, lang_code):
        self.current_lang = lang_code
        t = TRANSLATIONS[lang_code]

        if lang_code == 'ar':
            self.setLayoutDirection(Qt.RightToLeft)
        else:
            self.setLayoutDirection(Qt.LeftToRight)

        self.setWindowTitle(f"{t['app_title']} - [{self.username} : {self.role.upper()}]")
        self.btn_file.setText(t['file_btn'])
        self.act_entities.setText(t['entities_btn'])
        self.act_about.setText(t['about_btn'])
        if self.role == "admin":
            self.act_users.setText(t['users_btn'])
            self.act_logs.setText(t['logs_btn'])
            self.act_supplier_accounts.setText("Supplier Accounts")
            self.lbl_fyear.setText(t['fiscal_year_lbl'])
            self.btn_filter_year.setText(t['filter_btn'])
            self.btn_export.setText(t['export_excel_btn'])
            self.btn_export_cnt.setText(t['export_cnt_excel_btn'])
            self.btn_export_client.setText(t['export_client_excel_btn'])
            self.combo_fiscal_year.setItemText(0, t['all_years'])
            self.btn_nav_rep.setText(t['nav_rep'])
            self.btn_nav_cnt_rep.setText(t['nav_cnt_rep'])
            self.btn_nav_client_dossier.setText(t['nav_client_dossier'])

        self.btn_nav_pur.setText(t['nav_pur'])
        self.btn_nav_ship.setText(t['nav_ship'])
        self.btn_nav_pay.setText(t['nav_pay'])

        self.box_pur.setTitle(t['inv_box'])
        self.lbl_inv_num.setText(t['inv_num'])
        self.lbl_client1.setText(t['client'])
        self.lbl_supp.setText(t['supplier'])
        self.lbl_inv_tot.setText(t['inv_total'])
        self.lbl_inv_net.setText(t['inv_net'])
        self.lbl_inv_cbm.setText(t['inv_cbm'])
        if not self.editing_invoice_id:
            self.btn_save_pur.setText(t['save_inv_btn'])

        self.box_ship.setTitle(t['cnt_box'])
        self.lbl_cnt_num.setText(t['cnt_num'])
        self.lbl_client2.setText(t['client'])
        self.lbl_agent.setText(t['agent'])
        self.lbl_line.setText(t['shipping_line'])
        self.lbl_bol.setText(t['bol_num'])
        self.lbl_freight.setText(t['sea_freight'])
        self.lbl_customs.setText(t['customs'])
        self.lbl_comm.setText(t['commission'])
        self.lbl_total_cnt.setText(t['total_cnt_cost'])
        self.btn_save_ship.setText(t['save_cnt_btn'])

        self.sub_tabs.setTabText(0, t['subtab_client_pay'])
        self.sub_tabs.setTabText(1, t['subtab_expenses'])
        self.sub_tabs.setTabText(2, "Supplier Payments")
        self.lbl_pay_client.setText(t['client'])
        self.lbl_pay_amt.setText(t['amount'])
        self.lbl_pay_meth.setText(t['pay_method'])
        self.btn_attach.setText(t['attach_doc'])
        self.btn_save_pay.setText(t['save_pay_btn'])

        self.lbl_exp_cnt.setText(t['target_cnt'])
        self.lbl_exp_cat.setText(t['exp_category'])
        self.lbl_exp_notes.setText(t['notes'])
        self.btn_save_exp.setText(t['save_exp_btn'])

        # ---------------- تطبيق التلميحات الذكية المترجمة (Tooltips) ----------------
        self.pur_inv_num.setToolTip(t['tip_pur_inv_num'])
        self.pur_client.setToolTip(t['tip_pur_client'])
        self.pur_supplier.setToolTip(t['tip_pur_supplier'])
        self.pur_curr.setToolTip(t['tip_pur_curr'])
        self.pur_fx_rate.setToolTip(t['tip_pur_fx'])
        self.pur_total.setToolTip(t['tip_pur_total'])
        self.pur_net.setToolTip(t['tip_pur_net'])
        self.pur_cbm.setToolTip(t['tip_pur_cbm'])
        self.btn_import_pur_excel.setToolTip(t['tip_pur_import_excel'])
        self.btn_save_pur.setToolTip(t['tip_pur_save_btn'])

        self.ship_cnt_num.setToolTip(t['tip_cnt_num'])
        self.ship_client.setToolTip(t['tip_cnt_client'])
        self.ship_agent.setToolTip(t['tip_cnt_agent'])
        self.ship_line.setToolTip(t['tip_cnt_line'])
        self.ship_bol.setToolTip(t['tip_cnt_bol'])
        self.ship_status.setToolTip(t['tip_cnt_status'])
        self.ship_freight.setToolTip(t['tip_cnt_freight'])
        self.ship_freight_fx.setToolTip(t['tip_cnt_freight_fx'])
        self.ship_customs.setToolTip(t['tip_cnt_customs'])
        self.ship_comm.setToolTip(t['tip_cnt_comm'])
        self.ship_total_egp.setToolTip(t['tip_cnt_total_egp'])
        self.pay_client.setToolTip(t['tip_pay_client'])
        self.pay_client_amt.setToolTip(t['tip_pay_amt'])
        self.pay_fx.setToolTip(t['tip_pay_fx'])
        self.pay_client_method.setToolTip(t['tip_pay_meth'])
        self.btn_attach.setToolTip(t['tip_pay_attach'])

        self.exp_container.setToolTip(t['tip_exp_cnt'])
        self.exp_type.setToolTip(t['tip_exp_cat'])
        self.exp_amt.setToolTip(t['tip_exp_amt'])

        if self.role == "admin":
            self.load_admin_reports()
            self.load_container_manifest()
            self.load_client_dossier()

    def show_about_dialog(self):
        msg = f"<b>{APP_NAME} ERP System</b><br><br>Customs Clearance, Container Tracking & Financial Management.<br><br><b>Developer & Owner:</b> KO<br><b>Phone / WhatsApp:</b> 01023368006<br><br>{COPYRIGHT_TEXT}"
        QMessageBox.about(self, f"About {APP_NAME}", msg)

    def open_user_manager(self):
        dlg = UserManagerDialog(self.username)
        dlg.exec_()

    def open_entity_manager(self, default_type='CLIENT'):
        dlg = EntityManagerDialog(self.username, self.refresh_all_dropdowns, default_type)
        dlg.exec_()

    def refresh_all_dropdowns(self):
        for combo in [self.pur_client, self.ship_client, self.pay_client, self.pay_supplier]:
            combo.clear()
        self.pur_supplier.clear()
        self.ship_agent.clear()
        self.exp_container.clear()

        if self.role == "admin":
            self.combo_rep_cnt.clear()
            self.combo_dossier_client.clear()

        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()

        c.execute("SELECT id, name FROM entities WHERE entity_type = 'CLIENT'")
        for r in c.fetchall():
            self.pur_client.addItem(r[1], r[0])
            self.ship_client.addItem(r[1], r[0])
            self.pay_client.addItem(r[1], r[0])
            if self.role == "admin":
                self.combo_dossier_client.addItem(r[1], r[0])

        c.execute("SELECT id, name FROM entities WHERE entity_type = 'SUPPLIER'")
        for r in c.fetchall():
            self.pur_supplier.addItem(r[1], r[0])
            self.pay_supplier.addItem(r[1], r[0])

        c.execute("SELECT id, name FROM entities WHERE entity_type = 'SHIPPING_AGENT'")
        for r in c.fetchall():
            self.ship_agent.addItem(r[1], r[0])

        c.execute("SELECT id, container_num FROM containers")
        for r in c.fetchall():
            self.exp_container.addItem(f"Container: {r[1]}", r[0])
            if self.role == "admin":
                self.combo_rep_cnt.addItem(f"{r[1]}", r[0])

        conn.close()

    def pick_file(self, target_line_edit):
        path, _ = QFileDialog.getOpenFileName(self, "Select Document", "", "All Files (*.*)")
        if path:
            dest = os.path.join("attachments", os.path.basename(path))
            shutil.copy(path, dest)
            target_line_edit.setText(dest)


    def add_purchase_row(self):
        if self.pur_table.rowCount() >= 15:
            QMessageBox.warning(
                self, "Maximum Items",
                "Maximum 15 items per purchase invoice. Save this invoice and start a new invoice."
            )
            return
        row = self.pur_table.rowCount()
        self.pur_table.insertRow(row)
        for col in [5,7,9]:
            item = QTableWidgetItem("")
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.pur_table.setItem(row,col,item)

    def open_cargo_picker(self):
        client_id = self.ship_client.currentData()
        if not client_id:
            QMessageBox.warning(self, "Warning", "Please select a client first.")
            return
        dlg = ContainerCargoPickerDialog(client_id, self)
        if dlg.exec_() != QDialog.Accepted:
            return
        self.ship_items_table.setRowCount(0)
        for itm in dlg.selected_items:
            row = self.ship_items_table.rowCount()
            self.ship_items_table.insertRow(row)
            vals = [
                itm["invoice_num"], itm["code"], itm["name_ar"], itm["name_cn"],
                itm["cartons"], itm["pcs_per_carton"], itm["total_pcs"],
                "", itm["cbm"], itm["weight"]
            ]
            # fetch piece price from source invoice item
            conn = sqlite3.connect("import_enterprise.db")
            price_row = conn.execute("SELECT piece_price FROM invoice_items WHERE id=?", (itm["invoice_item_id"],)).fetchone()
            conn.close()
            vals[7] = price_row[0] if price_row else 0
            for col,val in enumerate(vals):
                item = QTableWidgetItem(str(val))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                if col == 0:
                    item.setData(Qt.UserRole, itm["invoice_item_id"])
                self.ship_items_table.setItem(row,col,item)
        self.ship_items_table.resizeRowsToContents()
        self.update_shipping_cbm_counter()

    def open_supplier_account(self):
        if self.role != "admin":
            return
        SupplierAccountDialog(self).exec_()

    def save_supplier_payment(self):
        if is_year_locked(self.current_year):
            QMessageBox.critical(self, "Fiscal Lock", f"Fiscal Year {self.current_year} is LOCKED!")
            return
        sid = self.pay_supplier.currentData()
        amt = float(self.supplier_pay_amt.text().replace(",","").strip() or 0)
        curr = self.supplier_pay_curr.currentText()
        fx = float(self.supplier_pay_fx.text().replace(",","").strip() or 1.0)
        amt_local = amt * fx
        method = self.supplier_pay_method.currentText()
        reason = self.supplier_pay_notes.text().strip() or "Supplier Payment"
        today_str = datetime.now().strftime("%Y-%m-%d")
        if not sid or amt <= 0:
            QMessageBox.warning(self, "Warning", "Select Supplier and enter a valid payment amount.")
            return
        conn = sqlite3.connect("import_enterprise.db")
        conn.execute("""INSERT INTO supplier_ledger
            (supplier_id, invoice_id, tx_type, amount, amount_local, currency, exchange_rate,
             supplier_invoice_num, reason, tx_date, method, created_by)
            VALUES (?, NULL, 'PAYMENT', ?, ?, ?, ?, NULL, ?, ?, ?, ?)""",
            (sid, amt, amt_local, curr, fx, reason, today_str, method, self.username))
        conn.commit()
        conn.close()
        record_log(self.username, "SUPPLIER_PAYMENT", f"Paid supplier '{self.pay_supplier.currentText()}': {amt:,.2f} {curr} ({amt_local:,.2f} EGP)")
        self.supplier_pay_amt.clear()
        self.supplier_pay_notes.clear()
        QMessageBox.information(self, "Success", "Supplier payment saved successfully.")


    # ---------------- حفظ وتوزيع مصاريف الحاوية ----------------
    def on_shipping_client_changed(self):
        """Generate the next container number for the selected client."""
        client_id = self.ship_client.currentData()
        if not client_id:
            self.ship_cnt_num.clear()
            return

        conn = sqlite3.connect("import_enterprise.db")
        try:
            c = conn.cursor()
            c.execute(
                "SELECT COUNT(*) FROM containers WHERE client_id = ?",
                (client_id,)
            )
            next_seq = (c.fetchone()[0] or 0) + 1

            # Keep the number unique globally while maintaining a separate
            # sequence for every client: C<client_id>-<client sequence>.
            container_num = f"C{int(client_id):03d}-{next_seq:04d}"

            # Protect against gaps/collisions if containers were deleted.
            while c.execute(
                "SELECT 1 FROM containers WHERE container_num = ?",
                (container_num,)
            ).fetchone():
                next_seq += 1
                container_num = f"C{int(client_id):03d}-{next_seq:04d}"

            self.ship_cnt_num.setText(container_num)
        finally:
            conn.close()

    def update_shipping_cbm_counter(self):
        """Update the visible CBM total for cargo currently added to the container."""
        total_cbm = 0.0
        for row in range(self.ship_items_table.rowCount()):
            try:
                value = self.ship_items_table.item(row, 8)
                if value:
                    total_cbm += float(value.text().replace(",", "").strip() or 0)
            except (ValueError, TypeError):
                pass

        self.lbl_ship_total_cbm.setText(
            f"Total Container CBM: {total_cbm:,.2f} CBM"
        )
        if total_cbm > 70:
            self.lbl_ship_total_cbm.setStyleSheet(
                "font-size: 15px; font-weight: bold; padding: 8px; color: #dc3545;"
            )
        else:
            self.lbl_ship_total_cbm.setStyleSheet(
                "font-size: 15px; font-weight: bold; padding: 8px; color: #198754;"
            )

    def save_container_data(self):
        if is_year_locked(self.current_year):
            QMessageBox.critical(self,"Fiscal Lock",f"Fiscal Year {self.current_year} is LOCKED! Cannot record container.")
            return
        c_num=self.ship_cnt_num.text().strip()
        client_id=self.ship_client.currentData()
        agent_id=self.ship_agent.currentData()
        line=self.ship_line.text().strip()
        bol=self.ship_bol.text().strip()
        freight_raw=self.ship_freight_curr.currentText()
        freight_curr="USD" if "USD" in freight_raw else ("RMB" if "RMB" in freight_raw else ("EUR" if "EUR" in freight_raw else "EGP"))
        try:
            freight_amt=float(self.ship_freight.text().replace(",","").strip() or 0)
            freight_fx=float(self.ship_freight_fx.text().replace(",","").strip() or 1)
            customs=float(self.ship_customs.text().replace(",","").strip() or 0)
            comm=float(self.ship_comm.text().replace(",","").strip() or 0)
        except ValueError:
            QMessageBox.warning(self,"Warning","Invalid freight/customs/commission amount.")
            return
        total_cost=freight_amt*freight_fx+customs+comm
        status=self.ship_status.currentText()
        today=datetime.now().strftime("%Y-%m-%d")
        if not c_num or not client_id:
            QMessageBox.warning(self,"Warning","Please select a Client. Container No is generated automatically.")
            return
        if self.ship_items_table.rowCount()==0:
            QMessageBox.warning(self,"Warning","Pull at least one unshipped invoice item into the container.")
            return

        payload=[]
        total_cbm=0.0
        for r in range(self.ship_items_table.rowCount()):
            iid=self.ship_items_table.item(r,0).data(Qt.UserRole) if self.ship_items_table.item(r,0) else None
            if not iid:
                QMessageBox.warning(self,"Warning","Invalid cargo row. Use Pull Unshipped Invoices.")
                return
            def txt(col):
                return self.ship_items_table.item(r,col).text().strip() if self.ship_items_table.item(r,col) else ""
            cartons=int(float(txt(4) or 0))
            ppc=max(1,int(float(txt(5) or 1)))
            total_pcs=int(float(txt(6) or cartons*ppc))
            price=float(txt(7).replace(",","") or 0)
            cbm=float(txt(8).replace(",","") or 0)
            weight=float(txt(9).replace(",","") or 0)
            payload.append((iid,txt(0),txt(1),txt(2),txt(3),cartons,ppc,total_pcs,price,total_pcs*price,cbm,weight))
            total_cbm+=cbm

        conn=sqlite3.connect("import_enterprise.db")
        try:
            cur=conn.cursor()
            for iid,*_ in payload:
                check=cur.execute("""SELECT client_invoices.client_id,COALESCE(invoice_items.shipped_status,0)
                                     FROM invoice_items JOIN client_invoices ON invoice_items.invoice_id=client_invoices.id
                                     WHERE invoice_items.id=?""",(iid,)).fetchone()
                if not check or check[0]!=client_id or check[1]:
                    raise ValueError("One or more selected invoice items are no longer available for shipment.")

            cur.execute("""INSERT INTO containers
                (container_num,client_id,agent_id,shipping_line,bill_of_lading,freight_currency,sea_freight,
                 freight_exchange_rate,customs_cost,commission,total_container_cost,status,arrival_date,currency,exchange_rate,fiscal_year)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,'EGP',1.0,?)""",
                (c_num,client_id,agent_id,line,bol,freight_curr,freight_amt,freight_fx,customs,comm,total_cost,status,today,self.current_year))
            container_id=cur.lastrowid

            for iid,inv_num,code,ar,cn,cartons,ppc,total_pcs,price,total_price,cbm,weight in payload:
                allocated=(cbm/total_cbm)*total_cost if total_cbm>0 and cbm>0 else 0.0
                landed=price+(allocated/total_pcs if total_pcs>0 else 0)
                cur.execute("""INSERT INTO container_items
                    (container_id,invoice_item_id,source_invoice_num,item_code,item_ar,item_cn,cartons,pcs_per_carton,total_pcs,
                     piece_price,total_price,cbm,total_cbm,weight,allocated_expense,landed_cost_unit)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (container_id,iid,inv_num,code,ar,cn,cartons,ppc,total_pcs,price,total_price,cbm,cbm,weight,allocated,landed))
                cur.execute("UPDATE invoice_items SET shipped_status=1 WHERE id=?",(iid,))

            cur.execute("""INSERT INTO ledger
                (tx_category,entity_id,container_id,currency,exchange_rate,amount,amount_local,tx_date,notes,fiscal_year,created_by)
                VALUES ('EXPENSE',?,?,?,?,?,?,?,?,?,?)""",
                (client_id,container_id,"EGP",1.0,total_cost,total_cost,today,
                 f"Container Cost: {c_num} ({status}) [Freight: {freight_amt}{freight_curr}, Customs: {customs:,.2f} EGP]",
                 self.current_year,self.username))

            if "Cleared" in status or "الإفراج" in status:
                cur.execute("UPDATE containers SET stocked_status=1 WHERE id=?",(container_id,))
                cur.execute("""INSERT INTO ledger
                    (tx_category,entity_id,container_id,currency,exchange_rate,amount,amount_local,tx_date,notes,fiscal_year,created_by)
                    VALUES ('WAREHOUSE_STOCK',?,?,?,?,?,?,?,?,?,?)""",
                    (client_id,container_id,"EGP",1.0,0,0,today,
                     f"Stock Transfer: Cargo for Container {c_num} Cleared & Stored",self.current_year,self.username))
            conn.commit()
            record_log(self.username,"SAVE_CONTAINER",f"Registered container '{c_num}' (Total: {total_cost:,.2f} EGP; {len(payload)} pulled invoice items)")
            QMessageBox.information(self,"Success",f"Container registered successfully.\nTotal Cost: {total_cost:,.2f} EGP allocated to cargo.")
            self.ship_cnt_num.clear()
            self.ship_items_table.setRowCount(0)
            self.update_shipping_cbm_counter()
            self.ship_freight.setText("0.0")
            self.ship_customs.setText("0.0")
            self.ship_comm.setText("0.0")
            self.ship_total_egp.setText("0.00 EGP")
            self.refresh_all_dropdowns()
            self.populate_fiscal_years()
        except ValueError as e:
            conn.rollback()
            QMessageBox.warning(self,"Cargo Validation",str(e))
        except sqlite3.IntegrityError:
            conn.rollback()
            QMessageBox.warning(self,"Error","Container No already exists!")
    def save_client_payment(self):
        if is_year_locked(self.current_year):
            QMessageBox.critical(self, "Fiscal Lock", f"Fiscal Year {self.current_year} is LOCKED!")
            return

        c_id = self.pay_client.currentData()
        c_name = self.pay_client.currentText()
        amt = float(self.pay_client_amt.text() or 0)
        curr = self.pay_curr.currentText()
        fx = float(self.pay_fx.text().strip() or 1.0)
        amt_local = amt * fx
        method = self.pay_client_method.currentText()
        doc = self.pay_client_doc.text()
        today_str = datetime.now().strftime("%Y-%m-%d")

        if not c_id or amt <= 0:
            QMessageBox.warning(self, "Warning", "Select Client and enter valid amount!")
            return

        conn = sqlite3.connect("import_enterprise.db")
        conn.execute("""INSERT INTO ledger (tx_category, entity_id, currency, exchange_rate, amount, amount_local, tx_date, method, doc_attachment, notes, fiscal_year, created_by)
                        VALUES ('CLIENT_PAY', ?, ?, ?, ?, ?, ?, ?, ?, 'Client Payment Deposit', ?, ?)""",
                     (c_id, curr, fx, amt, amt_local, today_str, method, doc, self.current_year, self.username))
        conn.commit()
        conn.close()

        record_log(self.username, "CLIENT_PAYMENT", f"Received {amt:,.2f} {curr} ({amt_local:,.2f} EGP) from '{c_name}' via {method}")
        self.pay_client_amt.clear()
        self.pay_client_doc.clear()
        QMessageBox.information(self, "Success", "Payment saved successfully.")
        self.populate_fiscal_years()

    def save_expense_data(self):
        if is_year_locked(self.current_year):
            QMessageBox.critical(self, "Fiscal Lock", f"Fiscal Year {self.current_year} is LOCKED!")
            return

        cnt_id = self.exp_container.currentData()
        cnt_text = self.exp_container.currentText()
        amt = float(self.exp_amt.text() or 0)
        curr = self.exp_curr.currentText()
        fx = float(self.exp_fx.text().strip() or 1.0)
        amt_local = amt * fx
        desc = f"{self.exp_type.currentText()} - {self.exp_notes.text()}"
        today_str = datetime.now().strftime("%Y-%m-%d")

        if not cnt_id or amt <= 0:
            QMessageBox.warning(self, "Warning", "Select Container and enter valid amount!")
            return

        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        c.execute("SELECT client_id FROM containers WHERE id = ?", (cnt_id,))
        row = c.fetchone()
        client_id = row[0] if row else None

        conn.execute("""INSERT INTO ledger (tx_category, entity_id, container_id, currency, exchange_rate, amount, amount_local, tx_date, notes, fiscal_year, created_by)
                        VALUES ('EXPENSE', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                     (client_id, cnt_id, curr, fx, amt, amt_local, today_str, desc, self.current_year, self.username))
        conn.commit()
        conn.close()

        record_log(self.username, "CONTAINER_EXPENSE", f"Recorded expense {amt:,.2f} {curr} on {cnt_text} ({desc})")
        self.exp_amt.clear()
        self.exp_notes.clear()
        QMessageBox.information(self, "Success", "Expense saved successfully.")
        self.populate_fiscal_years()

    def load_container_manifest(self):
        if self.role != "admin" or not hasattr(self, 'combo_rep_cnt'):
            return
        cnt_id = self.combo_rep_cnt.currentData()
        if not cnt_id:
            self.lbl_cnt_header.setText("Container Info: -")
            self.table_cnt_items.setRowCount(0)
            return

        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()

        c.execute("""
        SELECT c.container_num, e.name, c.shipping_line, c.bill_of_lading, c.arrival_date, c.status, c.total_container_cost
        FROM containers c
        LEFT JOIN entities e ON c.client_id = e.id
        WHERE c.id = ?
        """, (cnt_id,))
        cnt = c.fetchone()

        if cnt:
            info_text = f"Container: {cnt[0]} | Client: {cnt[1]} | Line: {cnt[2]} | B/L: {cnt[3]} | Arrival: {cnt[4]} | Status: {cnt[5]} | Total Cost: {cnt[6]:,.2f} EGP"
            self.lbl_cnt_header.setText(info_text)

        c.execute("""
        SELECT item_code, item_ar, item_cn, cartons, piece_price, cbm, weight, allocated_expense, landed_cost_unit 
        FROM container_items WHERE container_id = ?
        """, (cnt_id,))
        items = c.fetchall()
        conn.close()

        self.table_cnt_items.setRowCount(0)
        for idx, itm in enumerate(items):
            self.table_cnt_items.insertRow(idx)
            for col_idx, val in enumerate(itm):
                if col_idx in [4, 7, 8] and isinstance(val, (int, float)):
                    txt = f"{val:,.2f}"
                else:
                    txt = str(val if val is not None else "")
                self.table_cnt_items.setItem(idx, col_idx, QTableWidgetItem(txt))

    def mark_container_cleared(self):
        cnt_id = self.combo_rep_cnt.currentData()
        if not cnt_id:
            return
        
        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        c.execute("SELECT container_num, client_id FROM containers WHERE id = ?", (cnt_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return
            
        c_num, c_id = row
        today_str = datetime.now().strftime("%Y-%m-%d")
        
        c.execute("UPDATE containers SET status = 'تم الإفراج والمخزنة (Cleared & Stocked)', stocked_status = 1 WHERE id = ?", (cnt_id,))
        c.execute("""INSERT INTO ledger (tx_category, entity_id, container_id, currency, exchange_rate, amount, amount_local, tx_date, notes, fiscal_year, created_by)
                     VALUES ('WAREHOUSE_STOCK', ?, ?, 'EGP', 1.0, 0, 0, ?, ?, ?, ?)""",
                  (c_id, cnt_id, today_str, f"Stock Transfer: Cargo for Container {c_num} Cleared & Stored", self.current_year, self.username))
        conn.commit()
        conn.close()

        record_log(self.username, "WAREHOUSE_STOCKED", f"Marked container '{c_num}' as Cleared & Stocked")
        QMessageBox.information(self, "Success", "Container marked as Cleared & Stocked. Warehouse entry generated.")
        self.load_container_manifest()

    def open_attached_file(self, file_path):
        if file_path and os.path.exists(file_path):
            try:
                os.startfile(file_path)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open file: {e}")
        else:
            QMessageBox.information(self, "Notice", "No file attached or file path is invalid.")

    def load_client_dossier(self):
        if self.role != "admin" or not hasattr(self, 'combo_dossier_client'):
            return
        client_id = self.combo_dossier_client.currentData()
        if not client_id:
            self.lbl_client_stats.setText("Client Stats: Total Expenses: 0.00 | Total Paid: 0.00 | Balance: 0.00")
            self.table_client_cnts.setRowCount(0)
            self.table_client_pays.setRowCount(0)
            self.table_client_all_items.setRowCount(0)
            return

        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()

        c.execute("""
        SELECT 
            COALESCE(SUM(CASE WHEN tx_category = 'EXPENSE' THEN amount_local ELSE 0 END), 0) as total_exp,
            COALESCE(SUM(CASE WHEN tx_category = 'CLIENT_PAY' THEN amount_local ELSE 0 END), 0) as total_pay
        FROM ledger
        WHERE entity_id = ?
        """, (client_id,))
        stat = c.fetchone()
        exp = stat[0] if stat else 0
        pay = stat[1] if stat else 0
        bal = exp - pay
        bal_str = f"Debit (Owes {bal:,.2f} EGP)" if bal > 0 else (f"Credit (Overpaid {abs(bal):,.2f} EGP)" if bal < 0 else "Settled (0.00)")
        self.lbl_client_stats.setText(f"Financial Position: Total Expenses: {exp:,.2f} EGP | Total Paid: {pay:,.2f} EGP | Net Balance: {bal_str}")

        c.execute("""
        SELECT container_num, bill_of_lading, shipping_line, arrival_date, status, total_container_cost
        FROM containers
        WHERE client_id = ?
        ORDER BY id DESC
        """, (client_id,))
        cnts = c.fetchall()
        self.table_client_cnts.setRowCount(0)
        for idx, row in enumerate(cnts):
            self.table_client_cnts.insertRow(idx)
            for c_idx, val in enumerate(row):
                val_text = f"{val:,.2f}" if c_idx == 5 and isinstance(val, (int, float)) else str(val or "")
                self.table_client_cnts.setItem(idx, c_idx, QTableWidgetItem(val_text))

        c.execute("""
        SELECT tx_date, amount_local, method, notes, doc_attachment
        FROM ledger
        WHERE entity_id = ? AND tx_category = 'CLIENT_PAY'
        ORDER BY id DESC
        """, (client_id,))
        pays = c.fetchall()
        self.table_client_pays.setRowCount(0)
        for idx, row in enumerate(pays):
            self.table_client_pays.insertRow(idx)
            self.table_client_pays.setItem(idx, 0, QTableWidgetItem(str(row[0] or "")))
            self.table_client_pays.setItem(idx, 1, QTableWidgetItem(f"{row[1]:,.2f}"))
            self.table_client_pays.setItem(idx, 2, QTableWidgetItem(str(row[2] or "")))
            self.table_client_pays.setItem(idx, 3, QTableWidgetItem(str(row[3] or "")))

            doc_path = row[4]
            if doc_path and os.path.exists(doc_path):
                btn_view = QPushButton("👁️ View File")
                btn_view.setStyleSheet("background-color: #0d6efd; color: white; padding: 2px 6px; font-weight: bold;")
                btn_view.clicked.connect(lambda checked, p=doc_path: self.open_attached_file(p))
                self.table_client_pays.setCellWidget(idx, 4, btn_view)
            else:
                self.table_client_pays.setItem(idx, 4, QTableWidgetItem("No Attachment"))

        c.execute("""
        SELECT c.container_num, ci.item_code, ci.item_ar, ci.cartons,
               COALESCE(ci.total_cbm,ci.cbm,0), ci.weight, ci.allocated_expense, ci.landed_cost_unit
        FROM container_items ci
        JOIN containers c ON ci.container_id = c.id
        WHERE c.client_id = ?
        ORDER BY ci.id DESC
        """, (client_id,))
        all_items = c.fetchall()
        self.table_client_all_items.setRowCount(0)
        for idx, row in enumerate(all_items):
            self.table_client_all_items.insertRow(idx)
            for c_idx, val in enumerate(row):
                if c_idx in [6, 7] and isinstance(val, (int, float)):
                    txt = f"{val:,.2f}"
                else:
                    txt = str(val or "")
                self.table_client_all_items.setItem(idx, c_idx, QTableWidgetItem(txt))

        conn.close()

    def load_admin_reports(self):
        if self.role != "admin" or not hasattr(self, 'combo_fiscal_year'):
            return

        selected_year = self.combo_fiscal_year.currentData()
        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()

        if selected_year == "ALL" or selected_year is None:
            query = """
            SELECT e.name,
                   COALESCE(SUM(CASE WHEN l.tx_category = 'EXPENSE' THEN l.amount_local ELSE 0 END), 0) as total_expenses,
                   COALESCE(SUM(CASE WHEN l.tx_category = 'CLIENT_PAY' THEN l.amount_local ELSE 0 END), 0) as total_payments
            FROM entities e
            LEFT JOIN ledger l ON e.id = l.entity_id
            WHERE e.entity_type = 'CLIENT'
            GROUP BY e.id
            """
            c.execute(query)
        else:
            query = """
            SELECT e.name,
                   COALESCE(SUM(CASE WHEN l.tx_category = 'EXPENSE' AND l.fiscal_year = ? THEN l.amount_local ELSE 0 END), 0) as total_expenses,
                   COALESCE(SUM(CASE WHEN l.tx_category = 'CLIENT_PAY' AND l.fiscal_year = ? THEN l.amount_local ELSE 0 END), 0) as total_payments
            FROM entities e
            LEFT JOIN ledger l ON e.id = l.entity_id
            WHERE e.entity_type = 'CLIENT'
            GROUP BY e.id
            """
            c.execute(query, (selected_year, selected_year))

        rows = c.fetchall()
        conn.close()

        self.rep_table.setRowCount(0)
        for idx, (name, exp, pay) in enumerate(rows):
            bal = exp - pay
            status_en = "Debit (Owes Money)" if bal > 0 else ("Settled" if bal == 0 else "Credit (Overpaid)")
            self.rep_table.insertRow(idx)
            self.rep_table.setItem(idx, 0, QTableWidgetItem(name))
            self.rep_table.setItem(idx, 1, QTableWidgetItem(f"{exp:,.2f}"))
            self.rep_table.setItem(idx, 2, QTableWidgetItem(f"{pay:,.2f}"))
            self.rep_table.setItem(idx, 3, QTableWidgetItem(f"{abs(bal):,.2f}"))
            self.rep_table.setItem(idx, 4, QTableWidgetItem(status_en))

    def export_excel_english(self):
        selected_year = self.combo_fiscal_year.currentData()
        year_str = "All_Years" if selected_year == "ALL" or selected_year is None else f"FY_{selected_year}"
        default_filename = f"Kemt_Summary_{year_str}.xlsx"

        path, _ = QFileDialog.getSaveFileName(self, "Export Financial Summary", default_filename, "Excel (*.xlsx)")
        if not path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Summary {year_str}"
        ws.sheet_view.rightToLeft = False

        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        headers = ["Client Name", "Total Expenses (Local)", "Total Paid (Local)", "Remaining Balance", "Account Status"]
        ws.append(headers)

        for col_idx in range(1, 6):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()

        if selected_year == "ALL" or selected_year is None:
            c.execute("""
            SELECT e.name,
                   COALESCE(SUM(CASE WHEN l.tx_category = 'EXPENSE' THEN l.amount_local ELSE 0 END), 0),
                   COALESCE(SUM(CASE WHEN l.tx_category = 'CLIENT_PAY' THEN l.amount_local ELSE 0 END), 0)
            FROM entities e
            LEFT JOIN ledger l ON e.id = l.entity_id
            WHERE e.entity_type = 'CLIENT'
            GROUP BY e.id
            """)
        else:
            c.execute("""
            SELECT e.name,
                   COALESCE(SUM(CASE WHEN l.tx_category = 'EXPENSE' AND l.fiscal_year = ? THEN l.amount_local ELSE 0 END), 0),
                   COALESCE(SUM(CASE WHEN l.tx_category = 'CLIENT_PAY' AND l.fiscal_year = ? THEN l.amount_local ELSE 0 END), 0)
            FROM entities e
            LEFT JOIN ledger l ON e.id = l.entity_id
            WHERE e.entity_type = 'CLIENT'
            GROUP BY e.id
            """, (selected_year, selected_year))

        for row_idx, r in enumerate(c.fetchall(), start=2):
            bal = r[1] - r[2]
            st = "Debit (Owes Money)" if bal > 0 else ("Settled" if bal == 0 else "Credit (Overpaid)")
            vals = [r[0], r[1], r[2], abs(bal), st]
            for col_idx, val in enumerate(vals, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.border = thin_border
                if col_idx in [2, 3, 4]:
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
        conn.close()

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

        wb.save(path)
        record_log(self.username, "EXPORT_EXCEL", f"Exported Financial Summary: {os.path.basename(path)}")
        QMessageBox.information(self, "Export Complete", f"Summary report exported successfully to:\n{path}")

    def export_container_manifest_excel(self):
        cnt_id = self.combo_rep_cnt.currentData()
        cnt_num = self.combo_rep_cnt.currentText()
        if not cnt_id:
            QMessageBox.warning(self, "Warning", "Please select a container first!")
            return

        default_filename = f"Kemt_Container_{cnt_num}_Manifest.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Export Container Manifest", default_filename, "Excel (*.xlsx)")
        if not path:
            return

        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        c.execute("""
        SELECT c.container_num, e.name, c.shipping_line, c.bill_of_lading, c.arrival_date, c.status, c.total_container_cost
        FROM containers c
        LEFT JOIN entities e ON c.client_id = e.id
        WHERE c.id = ?
        """, (cnt_id,))
        cnt = c.fetchone()

        c.execute("""
        SELECT item_code, item_ar, item_cn, cartons, piece_price, cbm, weight, allocated_expense, landed_cost_unit 
        FROM container_items WHERE container_id = ?
        """, (cnt_id,))
        items = c.fetchall()
        conn.close()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Container {cnt_num}"
        ws.sheet_view.rightToLeft = False

        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        bold_font = Font(name="Calibri", size=11, bold=True)
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        ws.cell(row=1, column=1, value="Container No:").font = bold_font
        ws.cell(row=1, column=2, value=cnt[0])
        ws.cell(row=1, column=4, value="Client:").font = bold_font
        ws.cell(row=1, column=5, value=cnt[1])

        ws.cell(row=2, column=1, value="Shipping Line:").font = bold_font
        ws.cell(row=2, column=2, value=cnt[2])
        ws.cell(row=2, column=4, value="B/L Number:").font = bold_font
        ws.cell(row=2, column=5, value=cnt[3])

        ws.cell(row=3, column=1, value="Arrival Date:").font = bold_font
        ws.cell(row=3, column=2, value=cnt[4])
        ws.cell(row=3, column=4, value="Status:").font = bold_font
        ws.cell(row=3, column=5, value=cnt[5])

        ws.cell(row=4, column=1, value="Total Cost (Local):").font = bold_font
        ws.cell(row=4, column=2, value=f"{cnt[6]:,.2f} EGP").font = bold_font

        headers = ["Item Code", "Name (AR)", "Name (CN)", "Cartons", "Purchase Price", "CBM", "Weight (KG)", "Allocated Exp", "Landed Cost / Unit"]
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=6, column=col_idx, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row_idx, itm in enumerate(items, start=7):
            for col_idx, val in enumerate(itm, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.border = thin_border
                if col_idx in [4, 5, 6, 7, 8, 9]:
                    if col_idx in [5, 8, 9]:
                        cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right", vertical="center")

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

        wb.save(path)
        record_log(self.username, "EXPORT_EXCEL", f"Exported Container Manifest: {cnt_num}")
        QMessageBox.information(self, "Export Complete", f"Container manifest exported successfully to:\n{path}")

    def export_client_dossier_excel(self):
        client_id = self.combo_dossier_client.currentData()
        client_name = self.combo_dossier_client.currentText()
        if not client_id:
            QMessageBox.warning(self, "Warning", "Please select a client first!")
            return

        default_filename = f"Kemt_Client_{client_name}_Dossier.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Export Client Dossier", default_filename, "Excel (*.xlsx)")
        if not path:
            return

        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()

        c.execute("""
        SELECT 
            COALESCE(SUM(CASE WHEN tx_category = 'EXPENSE' THEN amount_local ELSE 0 END), 0),
            COALESCE(SUM(CASE WHEN tx_category = 'CLIENT_PAY' THEN amount_local ELSE 0 END), 0)
        FROM ledger WHERE entity_id = ?
        """, (client_id,))
        exp, pay = c.fetchone()
        bal = exp - pay

        c.execute("""
        SELECT container_num, bill_of_lading, shipping_line, arrival_date, status, total_container_cost
        FROM containers WHERE client_id = ? ORDER BY id DESC
        """, (client_id,))
        cnts = c.fetchall()

        c.execute("""
        SELECT tx_date, amount_local, method, notes, doc_attachment FROM ledger
        WHERE entity_id = ? AND tx_category = 'CLIENT_PAY' ORDER BY id DESC
        """, (client_id,))
        pays = c.fetchall()

        c.execute("""
        SELECT c.container_num, ci.item_code, ci.item_ar, ci.cartons,
               COALESCE(ci.total_cbm,ci.cbm,0), ci.weight, ci.allocated_expense, ci.landed_cost_unit
        FROM container_items ci
        JOIN containers c ON ci.container_id = c.id
        WHERE c.client_id = ? ORDER BY ci.id DESC
        """, (client_id,))
        items = c.fetchall()
        conn.close()

        wb = openpyxl.Workbook()
        ws_summary = wb.active
        ws_summary.title = "Financial & Containers"
        ws_summary.sheet_view.rightToLeft = False

        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        bold_font = Font(name="Calibri", size=11, bold=True)
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        ws_summary.cell(row=1, column=1, value="Client Dossier:").font = bold_font
        ws_summary.cell(row=1, column=2, value=client_name).font = bold_font
        ws_summary.cell(row=2, column=1, value="Total Costs (Local):").font = bold_font
        ws_summary.cell(row=2, column=2, value=f"{exp:,.2f} EGP")
        ws_summary.cell(row=3, column=1, value="Total Paid (Local):").font = bold_font
        ws_summary.cell(row=3, column=2, value=f"{pay:,.2f} EGP")
        ws_summary.cell(row=4, column=1, value="Remaining Balance:").font = bold_font
        ws_summary.cell(row=4, column=2, value=f"{bal:,.2f} EGP").font = bold_font

        ws_summary.cell(row=6, column=1, value="Containers Shipped:").font = bold_font
        cnt_headers = ["Container No", "B/L No", "Line", "Arrival Date", "Status", "Cost (Local)"]
        for col_idx, h in enumerate(cnt_headers, start=1):
            c_cell = ws_summary.cell(row=7, column=col_idx, value=h)
            c_cell.fill = header_fill
            c_cell.font = header_font

        for r_i, r_val in enumerate(cnts, start=8):
            for c_i, v in enumerate(r_val, start=1):
                cell = ws_summary.cell(row=r_i, column=c_i, value=v)
                cell.border = thin_border
                if c_i == 6:
                    cell.number_format = "#,##0.00"

        ws_pays = wb.create_sheet(title="Payment History")
        ws_pays.sheet_view.rightToLeft = False
        pay_headers = ["Date", "Amount (Local)", "Method", "Notes", "Receipt Document / Slip"]
        for col_idx, h in enumerate(pay_headers, start=1):
            c_cell = ws_pays.cell(row=1, column=col_idx, value=h)
            c_cell.fill = header_fill
            c_cell.font = header_font
            c_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws_pays.column_dimensions["E"].width = 24

        for r_i, r_val in enumerate(pays, start=2):
            ws_pays.cell(row=r_i, column=1, value=r_val[0]).border = thin_border
            amt_cell = ws_pays.cell(row=r_i, column=2, value=r_val[1])
            amt_cell.number_format = "#,##0.00"
            amt_cell.border = thin_border
            ws_pays.cell(row=r_i, column=3, value=r_val[2]).border = thin_border
            ws_pays.cell(row=r_i, column=4, value=r_val[3]).border = thin_border

            doc_path = r_val[4]
            doc_cell = ws_pays.cell(row=r_i, column=5)
            doc_cell.border = thin_border

            if doc_path and os.path.exists(doc_path):
                ext = os.path.splitext(doc_path)[1].lower()
                if ext in ['.png', '.jpg', '.jpeg', '.bmp']:
                    try:
                        xl_img = ExcelImage(doc_path)
                        xl_img.width = 110
                        xl_img.height = 70
                        ws_pays.add_image(xl_img, f"E{r_i}")
                        ws_pays.row_dimensions[r_i].height = 60
                    except Exception:
                        doc_cell.value = os.path.basename(doc_path)
                else:
                    doc_cell.value = f"Open File ({os.path.basename(doc_path)})"
                    doc_cell.hyperlink = os.path.abspath(doc_path)
                    doc_cell.font = Font(color="0000FF", underline="single")
            else:
                doc_cell.value = "No Attachment"

        ws_items = wb.create_sheet(title="Cargo & Landed Costs")
        ws_items.sheet_view.rightToLeft = False
        item_headers = ["Container No", "Item Code", "Name (AR)", "Cartons", "CBM", "Weight", "Allocated Exp", "Landed Cost / Unit"]
        for col_idx, h in enumerate(item_headers, start=1):
            c_cell = ws_items.cell(row=1, column=col_idx, value=h)
            c_cell.fill = header_fill
            c_cell.font = header_font

        for r_i, r_val in enumerate(items, start=2):
            for c_i, v in enumerate(r_val, start=1):
                cell = ws_items.cell(row=r_i, column=c_i, value=v)
                cell.border = thin_border
                if c_i in [7, 8]:
                    cell.number_format = "#,##0.00"

        for sheet in [ws_summary, ws_pays, ws_items]:
            for col in sheet.columns:
                col_letter = get_column_letter(col[0].column)
                if sheet == ws_pays and col_letter == "E":
                    continue
                max_len = max(len(str(cell.value or '')) for cell in col)
                sheet.column_dimensions[col_letter].width = max(max_len + 4, 15)

        wb.save(path)
        record_log(self.username, "EXPORT_EXCEL", f"Exported Client Dossier: {client_name}")
        QMessageBox.information(self, "Export Complete", f"Client dossier with embedded images exported to:\n{path}")

# -------------------------------------------------------------
# Dialogs for cargo selection and supplier accounts
# -------------------------------------------------------------

class ContainerCargoPickerDialog(QDialog):
    """Pick unshipped invoice items for a container."""
    def __init__(self, client_id, parent=None):
        super().__init__(parent)
        self.client_id = client_id
        self.selected_items = []
        self.setWindowTitle("📥 Pull Unshipped Invoices / اختيار بضاعة غير مشحونة")
        self.resize(1200, 620)

        layout = QVBoxLayout(self)
        top = QHBoxLayout()
        top.addWidget(QLabel("Select invoice / اختر الفاتورة:"))
        self.invoice_filter = QComboBox()
        self.invoice_filter.addItem("All Invoices", None)
        top.addWidget(self.invoice_filter)
        self.btn_invoice = QPushButton("☑ Select Invoice")
        self.btn_invoice.clicked.connect(self.select_current_invoice)
        top.addWidget(self.btn_invoice)
        self.btn_all = QPushButton("☑ Select All")
        self.btn_all.clicked.connect(self.select_all)
        top.addWidget(self.btn_all)
        self.btn_clear = QPushButton("☐ Clear")
        self.btn_clear.clicked.connect(self.clear_all)
        top.addWidget(self.btn_clear)
        top.addStretch()
        layout.addLayout(top)

        self.table = QTableWidget(0, 10)
        self.table.setHorizontalHeaderLabels([
            "Select", "Invoice No", "Item Code", "Description (AR)", "Description (CN)",
            "Cartons", "Pcs/Carton", "Total Pcs", "Total CBM", "Weight (KG)"
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.cellChanged.connect(self.update_counter)
        layout.addWidget(self.table)

        self.lbl_cbm = QLabel("Selected CBM: 0.00 / 70.00 CBM")
        self.lbl_cbm.setStyleSheet("font-weight: bold; padding: 6px;")
        layout.addWidget(self.lbl_cbm)

        buttons = QHBoxLayout()
        buttons.addStretch()
        ok = QPushButton("Pull Selected Cargo")
        ok.setStyleSheet("background-color:#0d6efd;color:white;font-weight:bold;padding:7px;")
        ok.clicked.connect(self.confirm_selection)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        buttons.addWidget(ok)
        buttons.addWidget(cancel)
        layout.addLayout(buttons)

        self._loading = False
        self.load_items()

    def load_items(self):
        conn = sqlite3.connect("import_enterprise.db")
        c = conn.cursor()
        c.execute("""
            SELECT ii.id, i.invoice_num, ii.item_code, ii.item_ar, ii.item_cn,
                   ii.cartons, COALESCE(ii.pcs_per_carton,1),
                   COALESCE(ii.total_pcs, ii.cartons*COALESCE(ii.pcs_per_carton,1)),
                   COALESCE(ii.total_cbm, ii.cbm, 0), ii.weight, i.id
            FROM invoice_items ii
            JOIN client_invoices i ON ii.invoice_id=i.id
            WHERE i.client_id=? AND COALESCE(ii.shipped_status,0)=0
            ORDER BY i.id DESC, ii.id
        """, (self.client_id,))
        rows = c.fetchall()
        conn.close()

        self._loading = True
        self.table.setRowCount(0)
        self.invoice_filter.blockSignals(True)
        self.invoice_filter.clear()
        self.invoice_filter.addItem("All Invoices", None)
        seen = set()
        for row in rows:
            inv_id = row[10]
            if inv_id not in seen:
                self.invoice_filter.addItem(row[1], inv_id)
                seen.add(inv_id)

            r = self.table.rowCount()
            self.table.insertRow(r)
            chk = QTableWidgetItem()
            chk.setCheckState(Qt.Unchecked)
            chk.setData(Qt.UserRole, row[0])
            self.table.setItem(r,0,chk)
            values = [row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9]]
            for col,val in enumerate(values,1):
                item = QTableWidgetItem(str(val if val is not None else ""))
                item.setData(Qt.UserRole, row[0])
                self.table.setItem(r,col,item)
        self.invoice_filter.blockSignals(False)
        self._loading = False
        self.update_counter()

    def selected_rows(self):
        ids = []
        for r in range(self.table.rowCount()):
            item = self.table.item(r,0)
            if item and item.checkState() == Qt.Checked:
                ids.append(r)
        return ids

    def select_all(self):
        self._loading = True
        for r in range(self.table.rowCount()):
            self.table.item(r,0).setCheckState(Qt.Checked)
        self._loading = False
        self.update_counter()

    def clear_all(self):
        self._loading = True
        for r in range(self.table.rowCount()):
            self.table.item(r,0).setCheckState(Qt.Unchecked)
        self._loading = False
        self.update_counter()

    def select_current_invoice(self):
        inv_id = self.invoice_filter.currentData()
        if not inv_id:
            self.select_all()
            return
        self._loading = True
        for r in range(self.table.rowCount()):
            item_id = self.table.item(r,0).data(Qt.UserRole)
            # invoice id is kept in the visible invoice column user data
            inv_text = self.table.item(r,1).text()
            target = self.invoice_filter.currentText()
            self.table.item(r,0).setCheckState(Qt.Checked if inv_text == target else Qt.Unchecked)
        self._loading = False
        self.update_counter()

    def update_counter(self, *args):
        if self._loading:
            return
        total = 0.0
        for r in self.selected_rows():
            try:
                total += float(self.table.item(r,8).text().replace(",",""))
            except Exception:
                pass
        self.lbl_cbm.setText(f"Selected CBM: {total:,.2f} / 70.00 CBM")
        if total > 70:
            self.lbl_cbm.setStyleSheet("font-weight:bold;padding:6px;color:#dc3545;")
        else:
            self.lbl_cbm.setStyleSheet("font-weight:bold;padding:6px;color:#198754;")

    def confirm_selection(self):
        rows = self.selected_rows()
        if not rows:
            QMessageBox.warning(self, "Warning", "Please select at least one invoice item.")
            return
        total = 0.0
        selected = []
        for r in rows:
            try:
                total += float(self.table.item(r,8).text().replace(",",""))
            except Exception:
                pass
            selected.append({
                "invoice_item_id": self.table.item(r,0).data(Qt.UserRole),
                "invoice_num": self.table.item(r,1).text(),
                "code": self.table.item(r,2).text(),
                "name_ar": self.table.item(r,3).text(),
                "name_cn": self.table.item(r,4).text(),
                "cartons": int(float(self.table.item(r,5).text() or 0)),
                "pcs_per_carton": int(float(self.table.item(r,6).text() or 1)),
                "total_pcs": int(float(self.table.item(r,7).text() or 0)),
                "cbm": float(self.table.item(r,8).text() or 0),
                "weight": float(self.table.item(r,9).text() or 0)
            })
        if total > 70:
            reply = QMessageBox.warning(
                self, "Container Capacity Warning",
                f"Selected cargo is {total:,.2f} CBM, above the standard 70 CBM capacity. Continue?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return
        self.selected_items = selected
        self.accept()


class SupplierAccountDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Supplier Account / حساب المورد")
        self.resize(650, 430)
        layout = QVBoxLayout(self)
        form = QHBoxLayout()
        form.addWidget(QLabel("Supplier:"))
        self.combo = QComboBox()
        form.addWidget(self.combo)
        layout.addLayout(form)

        self.lbl_invoices = QLabel("Total Invoices: 0.00 EGP")
        self.lbl_paid = QLabel("Total Paid: 0.00 EGP")
        self.lbl_balance = QLabel("Remaining Balance: 0.00 EGP")
        for w in [self.lbl_invoices,self.lbl_paid,self.lbl_balance]:
            w.setStyleSheet("font-size:15px;font-weight:bold;padding:8px;")
            layout.addWidget(w)

        self.table = QTableWidget(0,6)
        self.table.setHorizontalHeaderLabels(["Date","Type","Amount","Local (EGP)","Invoice","Reason"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)
        self.combo.currentIndexChanged.connect(self.load_account)
        self.load_suppliers()

    def load_suppliers(self):
        conn=sqlite3.connect("import_enterprise.db")
        rows=conn.execute("SELECT id,name FROM entities WHERE entity_type='SUPPLIER' ORDER BY name").fetchall()
        conn.close()
        self.combo.clear()
        for r in rows:
            self.combo.addItem(r[1],r[0])
        self.load_account()

    def load_account(self):
        sid=self.combo.currentData()
        if not sid:
            self.lbl_invoices.setText("Total Invoices: 0.00 EGP")
            self.lbl_paid.setText("Total Paid: 0.00 EGP")
            self.lbl_balance.setText("Remaining Balance: 0.00 EGP")
            self.table.setRowCount(0)
            return
        conn=sqlite3.connect("import_enterprise.db")
        c=conn.cursor()
        c.execute("""SELECT
                     COALESCE(SUM(CASE WHEN tx_type='INVOICE' THEN amount_local ELSE 0 END),0),
                     COALESCE(SUM(CASE WHEN tx_type='PAYMENT' THEN amount_local ELSE 0 END),0)
                     FROM supplier_ledger WHERE supplier_id=?""",(sid,))
        inv,paid=c.fetchone()
        c.execute("""SELECT tx_date,tx_type,amount,amount_local,supplier_invoice_num,reason
                     FROM supplier_ledger WHERE supplier_id=? ORDER BY id DESC""",(sid,))
        rows=c.fetchall()
        conn.close()
        self.lbl_invoices.setText(f"Total Invoices: {inv:,.2f} EGP")
        self.lbl_paid.setText(f"Total Paid: {paid:,.2f} EGP")
        self.lbl_balance.setText(f"Remaining Balance: {inv-paid:,.2f} EGP")
        self.table.setRowCount(0)
        for r,row in enumerate(rows):
            self.table.insertRow(r)
            for col,val in enumerate(row):
                self.table.setItem(r,col,QTableWidgetItem(str(val if val is not None else "")))


# -------------------------------------------------------------
# تشغيل التطبيق
# -------------------------------------------------------------
if __name__ == "__main__":
    init_database()
    app = QApplication(sys.argv)

    main_app = None
    def run_app(u, r):
        global main_app
        main_app = MainEnterpriseApp(u, r)
        main_app.show()

    login = LoginDialog(run_app)
    login.show()
    sys.exit(app.exec_())
