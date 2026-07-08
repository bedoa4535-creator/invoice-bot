#!/usr/bin/env python3
import os, re, logging, json
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

BOT_TOKEN = os.environ.get("BOT_TOKEN")
ADMIN_ID = int(os.environ.get("ADMIN_ID", "0"))

logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)

valid_codes = {}
authorized_users = {}
user_templates = {}
CODES_FILE = "codes.json"

def save_codes():
    with open(CODES_FILE, 'w') as f:
        json.dump({'codes': valid_codes, 'users': {str(k): v for k, v in authorized_users.items()}}, f)

def load_codes():
    global valid_codes, authorized_users
    if os.path.exists(CODES_FILE):
        with open(CODES_FILE, 'r') as f:
            data = json.load(f)
            valid_codes = data.get('codes', {})
            authorized_users = {int(k): v for k, v in data.get('users', {}).items()}

load_codes()

WAITING_CODE = 1
WAITING_TEMPLATE_CHOICE = 2

# ====== حقول كل شكل ======
DRAGONS_FIELDS = ['الاسم', 'الرقم', 'العنوان', 'اللون', 'العرض', 'المقاس', 'العدد', 'الاجمالي']
NF_FIELDS = ['اسم العميل', 'اكونت العميل', 'رقم الفاتورة', 'اسم الصفحة', 'اسم IT', 'العنوان', 'الرقم', 'الألوان', 'العدد', 'المقاس', 'نوع المنتج', 'سعر المنتج', 'مصاريف الشحن', 'الاجمالي']
MM_FIELDS = ['اسم الاكونت', 'IT', 'الاسم', 'العنوان', 'موبايل 1', 'موبايل 2', 'المنتج', 'الالوان', 'المقاس', 'الكمية', 'السعر', 'الشحن', 'الاجمالي']

TEMPLATES = {
    'dragon': DRAGONS_FIELDS,
    'fashion': NF_FIELDS,
    'mm': MM_FIELDS,
}

# ====== مساعد Word ======
def set_border(cell, color='000000', sz='4'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for side in ['top', 'left', 'bottom', 'right']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), sz)
        b.set(qn('w:color'), color)
        tcPr.append(b)

def set_bg(cell, color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def cell_text(cell, text, bold=True, size=9, align=WD_ALIGN_PARAGRAPH.RIGHT, color=None):
    para = cell.paragraphs[0]
    para.alignment = align
    run = para.add_run(str(text or ''))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = 'Arial'
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    pPr = para._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    pPr.append(bidi)

# ====== Dragon.s ======
def make_dragons_invoice(doc, inv):
    NOTE = 'في حالة عدم استلام الاوردر يتم دفع مصريف الشحن كامله للمندوب.'
    table = doc.add_table(rows=0, cols=4)
    table.style = 'Table Grid'
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[3])
    set_border(cell); set_bg(cell, 'D9D9D9')
    cell_text(cell, 'Dragon.s', bold=True, size=12)
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[3])
    set_border(cell); cell_text(cell, 'التاريخ/ رقم الاوردر/', bold=True)
    for label, field in [('الاسم:', 'الاسم'), ('الرقم', 'الرقم'), ('العنوان', 'العنوان')]:
        row = table.add_row()
        set_border(row.cells[0]); set_bg(row.cells[0], 'D9D9D9'); cell_text(row.cells[0], label)
        cell = row.cells[1].merge(row.cells[3]); set_border(cell); cell_text(cell, inv.get(field, ''))
    row = table.add_row()
    set_border(row.cells[0]); set_bg(row.cells[0], 'D9D9D9'); cell_text(row.cells[0], 'اللون:')
    set_border(row.cells[1]); cell_text(row.cells[1], inv.get('اللون', ''))
    set_border(row.cells[2]); set_bg(row.cells[2], 'D9D9D9'); cell_text(row.cells[2], 'العرض:')
    set_border(row.cells[3]); cell_text(row.cells[3], inv.get('العرض', ''))
    row = table.add_row()
    set_border(row.cells[0]); set_bg(row.cells[0], 'D9D9D9'); cell_text(row.cells[0], 'المقاس')
    set_border(row.cells[1]); cell_text(row.cells[1], inv.get('المقاس', ''))
    set_border(row.cells[2]); set_bg(row.cells[2], 'D9D9D9'); cell_text(row.cells[2], 'العدد:')
    set_border(row.cells[3]); cell_text(row.cells[3], inv.get('العدد', ''))
    row = table.add_row()
    set_border(row.cells[0]); set_bg(row.cells[0], 'D9D9D9'); cell_text(row.cells[0], 'الاجمالي')
    set_border(row.cells[1]); cell_text(row.cells[1], inv.get('الاجمالي', ''))
    set_border(row.cells[2]); set_bg(row.cells[2], 'D9D9D9'); cell_text(row.cells[2], 'IT :')
    set_border(row.cells[3]); cell_text(row.cells[3], '')
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[3]); set_border(cell)
    cell_text(cell, NOTE, bold=True, size=9)

# ====== New Fashion ======
def make_nf_invoice(doc, inv):
    NOTE = 'عزيزي العميل يحق لك قبل الاستلام معاينة المنتج والتأكد من المقاس ومن جودة المنتج  في حاله عدم الاستلام برجاء دفع مصاريف الشحن للمندوب    الاستبدال حد اقصي ثلاث ايام من تاريخ التسليم مع دفع مصاريف الشحن مره اخري ولا يوجد استرجاع'
    today = datetime.now().strftime('%d / %m / %Y')
    table = doc.add_table(rows=0, cols=6)
    table.style = 'Table Grid'
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[5])
    set_border(cell); set_bg(cell, '1F3864')
    cell_text(cell, '--NEW FASHION COMPANY--', bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, color='FFFFFF')
    row = table.add_row()
    set_border(row.cells[0]); cell_text(row.cells[0], inv.get('رقم الفاتورة', ''))
    set_border(row.cells[1]); set_bg(row.cells[1], 'D9D9D9'); cell_text(row.cells[1], 'رقم الفاتورة')
    set_border(row.cells[2]); cell_text(row.cells[2], inv.get('اسم IT', ''))
    set_border(row.cells[3]); set_bg(row.cells[3], 'D9D9D9'); cell_text(row.cells[3], 'اسم الـ IT')
    set_border(row.cells[4]); cell_text(row.cells[4], inv.get('اسم العميل', ''))
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'اسم العميل')
    row = table.add_row()
    set_border(row.cells[0]); cell_text(row.cells[0], today)
    set_border(row.cells[1]); set_bg(row.cells[1], 'D9D9D9'); cell_text(row.cells[1], 'تاريخ الفاتورة')
    set_border(row.cells[2]); cell_text(row.cells[2], inv.get('اسم الصفحة', ''))
    set_border(row.cells[3]); set_bg(row.cells[3], 'D9D9D9'); cell_text(row.cells[3], 'اسم الصفحة')
    set_border(row.cells[4]); cell_text(row.cells[4], inv.get('اكونت العميل', ''))
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'اكونت العميل')
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[4]); set_border(cell); cell_text(cell, inv.get('العنوان', ''))
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'العنوان')
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[4]); set_border(cell); cell_text(cell, inv.get('الرقم', ''))
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'الرقم')
    row = table.add_row()
    set_border(row.cells[0]); cell_text(row.cells[0], inv.get('الألوان', ''))
    cell = row.cells[1].merge(row.cells[2]); set_border(cell); set_bg(cell, 'D9D9D9'); cell_text(cell, 'الألوان')
    set_border(row.cells[3]); cell_text(row.cells[3], inv.get('العدد', ''))
    set_border(row.cells[4]); set_bg(row.cells[4], 'D9D9D9'); cell_text(row.cells[4], 'العدد')
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9')
    row = table.add_row()
    set_border(row.cells[0]); cell_text(row.cells[0], inv.get('المقاس', ''))
    cell = row.cells[1].merge(row.cells[2]); set_border(cell); set_bg(cell, 'D9D9D9'); cell_text(cell, 'المقاس')
    set_border(row.cells[3]); cell_text(row.cells[3], inv.get('نوع المنتج', ''))
    set_border(row.cells[4]); set_bg(row.cells[4], 'D9D9D9'); cell_text(row.cells[4], 'نوع المنتج')
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9')
    row = table.add_row()
    set_border(row.cells[0]); cell_text(row.cells[0], inv.get('الاجمالي', ''))
    set_border(row.cells[1]); set_bg(row.cells[1], 'D9D9D9'); cell_text(row.cells[1], 'الاجمالي')
    set_border(row.cells[2]); cell_text(row.cells[2], inv.get('مصاريف الشحن', ''))
    set_border(row.cells[3]); set_bg(row.cells[3], 'D9D9D9'); cell_text(row.cells[3], 'مصاريف الشحن')
    set_border(row.cells[4]); cell_text(row.cells[4], inv.get('سعر المنتج', ''))
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'سعر المنتج')
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[5]); set_border(cell); set_bg(cell, 'FFF2CC')
    cell_text(cell, NOTE, bold=False, size=8)

# ====== M&M ======
def make_mm_invoice(doc, inv):
    NOTE = 'يجب دفع خدمة مصاريف الشحن لمندوب شركة الشحن فى حالة الشراء او عدم الشراء نظرا لتكلفة مصاريف الانتقال'
    today = datetime.now().strftime('%d / %m / %Y')
    table = doc.add_table(rows=0, cols=6)
    table.style = 'Table Grid'

    # هيدر M&M
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[5])
    set_border(cell); set_bg(cell, '000000')
    cell_text(cell, 'M&M', bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, color='FFFFFF')

    # التاريخ / رقم الفاتورة / اسم الاكونت / IT
    row = table.add_row()
    set_border(row.cells[0]); cell_text(row.cells[0], today)
    set_border(row.cells[1]); set_bg(row.cells[1], 'D9D9D9'); cell_text(row.cells[1], 'التاريخ')
    set_border(row.cells[2]); cell_text(row.cells[2], inv.get('اسم الاكونت', ''))
    set_border(row.cells[3]); set_bg(row.cells[3], 'D9D9D9'); cell_text(row.cells[3], 'اسم الاكونت')
    set_border(row.cells[4]); cell_text(row.cells[4], inv.get('IT', ''))
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'IT')

    # الاسم
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[4]); set_border(cell); cell_text(cell, inv.get('الاسم', ''))
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'الاسم')

    # العنوان
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[4]); set_border(cell); cell_text(cell, inv.get('العنوان', ''))
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'العنوان بالتفصيل')

    # موبايل 1
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[4]); set_border(cell); cell_text(cell, inv.get('موبايل 1', ''))
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'رقم الموبايل 1')

    # موبايل 2
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[4]); set_border(cell); cell_text(cell, inv.get('موبايل 2', ''))
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'رقم الموبايل 2')

    # المنتج
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[4]); set_border(cell); cell_text(cell, inv.get('المنتج', ''))
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'المنتج 1')

    # الالوان / المقاس
    row = table.add_row()
    set_border(row.cells[0]); cell_text(row.cells[0], inv.get('الالوان', ''))
    cell = row.cells[1].merge(row.cells[2]); set_border(cell); set_bg(cell, 'D9D9D9'); cell_text(cell, 'الالوان')
    set_border(row.cells[3]); cell_text(row.cells[3], inv.get('المقاس', ''))
    set_border(row.cells[4]); set_bg(row.cells[4], 'D9D9D9'); cell_text(row.cells[4], 'المقاس')
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9')

    # الكمية / السعر / الشحن / الاجمالي
    row = table.add_row()
    set_border(row.cells[0]); cell_text(row.cells[0], inv.get('الكمية', ''))
    set_border(row.cells[1]); set_bg(row.cells[1], 'D9D9D9'); cell_text(row.cells[1], 'الكمية')
    set_border(row.cells[2]); cell_text(row.cells[2], inv.get('السعر', ''))
    set_border(row.cells[3]); set_bg(row.cells[3], 'D9D9D9'); cell_text(row.cells[3], 'السعر')
    set_border(row.cells[4]); cell_text(row.cells[4], inv.get('الشحن', ''))
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'الشحن')

    # الاجمالي
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[4]); set_border(cell); cell_text(cell, inv.get('الاجمالي', ''), bold=True, size=11)
    set_border(row.cells[5]); set_bg(row.cells[5], 'D9D9D9'); cell_text(row.cells[5], 'الاجمالي')

    # ملاحظة
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[5]); set_border(cell); set_bg(cell, 'FFF2CC')
    cell_text(cell, NOTE, bold=False, size=8)

def parse_invoices(text, fields):
    text = re.sub(r'^/invoice\s*', '', text.strip(), flags=re.IGNORECASE)
    blocks = re.split(r'\n\s*---+\s*\n', text)
    invoices = []
    for block in blocks:
        block = block.strip()
        if not block: continue
        inv = {}
        for line in block.splitlines():
            line = line.strip()
            for field in fields:
                if line.startswith(field + ':'):
                    inv[field] = line[len(field)+1:].strip()
                    break
        if inv: invoices.append(inv)
    return invoices

def create_word(invoices, template_type):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21); section.page_height = Cm(29.7)
    section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Cm(1)
    for i, inv in enumerate(invoices):
        if template_type == 'dragon': make_dragons_invoice(doc, inv)
        elif template_type == 'fashion': make_nf_invoice(doc, inv)
        elif template_type == 'mm': make_mm_invoice(doc, inv)
        if i < len(invoices) - 1: doc.add_paragraph()
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ====== رسائل المساعدة ======
HELP_MESSAGES = {
    'dragon': (
        "✅ اخترت شكل Dragon.s!\n\n"
        "ابعت /invoice مع البيانات:\n\n"
        "/invoice\n"
        "الاسم: محمد أحمد\n"
        "الرقم: 01012345678\n"
        "العنوان: القاهرة\n"
        "اللون: ازرق\n"
        "العرض: ترنج مشجر\n"
        "المقاس: لارج\n"
        "العدد: 1\n"
        "الاجمالي: 300"
    ),
    'fashion': (
        "✅ اخترت شكل New Fashion!\n\n"
        "ابعت /invoice مع البيانات:\n\n"
        "/invoice\n"
        "اسم العميل: محمد أحمد\n"
        "اكونت العميل: 01012345678\n"
        "رقم الفاتورة: 001\n"
        "اسم الصفحة: نيو فاشون\n"
        "اسم IT: أحمد\n"
        "العنوان: القاهرة\n"
        "الرقم: 01012345678\n"
        "الألوان: أحمر\n"
        "العدد: 2\n"
        "المقاس: XL\n"
        "نوع المنتج: تيشيرت\n"
        "سعر المنتج: 200\n"
        "مصاريف الشحن: 30\n"
        "الاجمالي: 430"
    ),
    'mm': (
        "✅ اخترت شكل M&M!\n\n"
        "ابعت /invoice مع البيانات:\n\n"
        "/invoice\n"
        "اسم الاكونت: محمد أحمد\n"
        "IT: سارة\n"
        "الاسم: محمود علي\n"
        "العنوان: القاهرة، مدينة نصر\n"
        "موبايل 1: 01012345678\n"
        "موبايل 2: 01098765432\n"
        "المنتج: بنطلون جينز\n"
        "الالوان: أزرق\n"
        "المقاس: XL\n"
        "الكمية: 2\n"
        "السعر: 350\n"
        "الشحن: 40\n"
        "الاجمالي: 390"
    )
}

# ====== هاندلرز ======
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id in authorized_users:
        await update.message.reply_text(
            "أهلاً! 👋\n\n"
            "ابعتلي اسم شكل الفاتورة بتاعتك:\n"
            "• *dragon*\n• *fashion*\n• *mm*",
            parse_mode='Markdown'
        )
        return WAITING_TEMPLATE_CHOICE
    else:
        await update.message.reply_text("أهلاً! 👋\n\n🔐 البوت ده خاص\nابعتلي كود الدخول بتاعك:")
        return WAITING_CODE

async def check_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    code = update.message.text.strip()
    if code in valid_codes:
        authorized_users[user_id] = code
        save_codes()
        await update.message.reply_text(
            "✅ الكود صح!\n\n"
            "ابعتلي اسم شكل الفاتورة بتاعتك:\n"
            "• *dragon*\n• *fashion*\n• *mm*",
            parse_mode='Markdown'
        )
        return WAITING_TEMPLATE_CHOICE
    else:
        await update.message.reply_text("❌ الكود غلط! جرب تاني")
        return WAITING_CODE

async def choose_template(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    choice = update.message.text.strip().lower()
    if choice in TEMPLATES:
        user_templates[user_id] = choice
        await update.message.reply_text(HELP_MESSAGES[choice])
        return ConversationHandler.END
    else:
        await update.message.reply_text(
            "❌ مش عارف الشكل ده!\n\n"
            "اكتب واحد من دول:\n• *dragon*\n• *fashion*\n• *mm*",
            parse_mode='Markdown'
        )
        return WAITING_TEMPLATE_CHOICE

async def invoice_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text or ""
    if user_id not in authorized_users:
        await update.message.reply_text("❌ محتاج كود دخول! ابعت /start")
        return
    template_type = user_templates.get(user_id)
    if not template_type:
        await update.message.reply_text("❌ ابعت /start الأول عشان تختار شكل الفاتورة")
        return
    if len(text.strip()) < 12:
        await update.message.reply_text(HELP_MESSAGES.get(template_type, "ابعت البيانات مع /invoice"))
        return
    await update.message.reply_text("⏳ بشتغل على الفواتير...")
    fields = TEMPLATES[template_type]
    invoices = parse_invoices(text, fields)
    if not invoices:
        await update.message.reply_text("❌ مش لاقي فواتير!")
        return
    try:
        buf = create_word(invoices, template_type)
        await update.message.reply_document(
            document=buf,
            filename=f"فواتير_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            caption=f"✅ تم إنشاء *{len(invoices)} فاتورة* بنجاح!",
            parse_mode='Markdown'
        )
    except Exception as e:
        await update.message.reply_text(f"❌ خطأ: {e}")

async def addcode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID: return
    if not context.args: await update.message.reply_text("استخدام: /addcode الكود"); return
    code = context.args[0]
    valid_codes[code] = datetime.now().strftime('%Y-%m-%d')
    save_codes()
    await update.message.reply_text(f"✅ تم إضافة الكود: `{code}`", parse_mode='Markdown')

async def removecode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID: return
    if not context.args: await update.message.reply_text("استخدام: /removecode الكود"); return
    code = context.args[0]
    if code in valid_codes:
        del valid_codes[code]
        to_remove = [uid for uid, c in authorized_users.items() if c == code]
        for uid in to_remove: del authorized_users[uid]
        save_codes()
        await update.message.reply_text(f"✅ تم حذف الكود: `{code}`", parse_mode='Markdown')
    else:
        await update.message.reply_text("❌ الكود مش موجود")

async def listcodes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID: return
    if not valid_codes: await update.message.reply_text("مفيش أكواد"); return
    text = "📋 *الأكواد:*\n\n"
    for code, date in valid_codes.items():
        users = [uid for uid, c in authorized_users.items() if c == code]
        text += f"• `{code}` - {date} - {len(users)} مستخدم\n"
    await update.message.reply_text(text, parse_mode='Markdown')

async def listusers(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID: return
    if not authorized_users: await update.message.reply_text("مفيش مستخدمين"); return
    text = f"👥 *المستخدمين: {len(authorized_users)}*\n\n"
    for uid, code in authorized_users.items():
        template = user_templates.get(uid, 'مش محدد')
        text += f"• ID: `{uid}` - كود: `{code}` - شكل: {template}\n"
    await update.message.reply_text(text, parse_mode='Markdown')

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in authorized_users:
        await update.message.reply_text("ابعت /start للبدء")
    else:
        await update.message.reply_text("ابعت /invoice لعمل فواتير أو /start لتغيير الشكل 🙂")

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            WAITING_CODE: [MessageHandler(filters.TEXT & ~filters.COMMAND, check_code)],
            WAITING_TEMPLATE_CHOICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, choose_template)],
        },
        fallbacks=[]
    )
    app.add_handler(conv)
    app.add_handler(CommandHandler("invoice", invoice_cmd))
    app.add_handler(CommandHandler("addcode", addcode))
    app.add_handler(CommandHandler("removecode", removecode))
    app.add_handler(CommandHandler("listcodes", listcodes))
    app.add_handler(CommandHandler("listusers", listusers))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("✅ البوت شغال!")
    app.run_polling()

if __name__ == '__main__':
    main()
