#!/usr/bin/env python3
import os, re, logging, json
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

BOT_TOKEN = os.environ.get("BOT_TOKEN")
ADMIN_ID = int(os.environ.get("ADMIN_ID", "909630778"))  # ← هتحط ID بتاعك هنا

logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)

# ====== داتا ======
valid_codes = {}      # الأكواد المتاحة: {code: description}
authorized_users = {} # اليوزرز المصرح ليهم: {user_id: code}
user_templates = {}   # templates: {user_id: bytes}
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

# ====== حالات ======
WAITING_CODE = 1
WAITING_TEMPLATE = 2

FIELDS = ['الاسم', 'الرقم', 'العنوان', 'اللون', 'العرض', 'المقاس', 'العدد', 'الاجمالي']
NOTE = 'في حالة عدم استلام الاوردر يتم دفع مصريف الشحن كامله للمندوب.'

# ====== Word ======
def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:color'), '000000')
        tcPr.append(border)

def set_cell_bg(cell, color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def add_cell_text(cell, text, bold=True, size=10):
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = para.add_run(str(text or ''))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = 'Arial'
    pPr = para._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    pPr.append(bidi)

def make_invoice_table(doc, inv):
    table = doc.add_table(rows=0, cols=4)
    table.style = 'Table Grid'

    row = table.add_row()
    cell = row.cells[0].merge(row.cells[3])
    set_cell_border(cell); set_cell_bg(cell, 'D9D9D9')
    add_cell_text(cell, 'Dragon.s', bold=True, size=12)

    row = table.add_row()
    cell = row.cells[0].merge(row.cells[3])
    set_cell_border(cell)
    add_cell_text(cell, 'التاريخ/ رقم الاوردر/', bold=True)

    for label, field, merge in [('الاسم:', 'الاسم', True), ('الرقم', 'الرقم', True), ('العنوان', 'العنوان', True)]:
        row = table.add_row()
        set_cell_border(row.cells[0]); set_cell_bg(row.cells[0], 'D9D9D9')
        add_cell_text(row.cells[0], label)
        cell = row.cells[1].merge(row.cells[3])
        set_cell_border(cell)
        add_cell_text(cell, inv.get(field, ''))

    row = table.add_row()
    for i, (label, field, gray) in enumerate([('اللون:', 'اللون', True), (None, 'اللون', False), ('العرض:', 'العرض', True), (None, 'العرض', False)]):
        set_cell_border(row.cells[i])
        if i == 0: set_cell_bg(row.cells[i], 'D9D9D9'); add_cell_text(row.cells[i], 'اللون:')
        elif i == 1: add_cell_text(row.cells[i], inv.get('اللون', ''))
        elif i == 2: set_cell_bg(row.cells[i], 'D9D9D9'); add_cell_text(row.cells[i], 'العرض:')
        elif i == 3: add_cell_text(row.cells[i], inv.get('العرض', ''))

    row = table.add_row()
    set_cell_border(row.cells[0]); set_cell_bg(row.cells[0], 'D9D9D9'); add_cell_text(row.cells[0], 'المقاس')
    set_cell_border(row.cells[1]); add_cell_text(row.cells[1], inv.get('المقاس', ''))
    set_cell_border(row.cells[2]); set_cell_bg(row.cells[2], 'D9D9D9'); add_cell_text(row.cells[2], 'العدد:')
    set_cell_border(row.cells[3]); add_cell_text(row.cells[3], inv.get('العدد', ''))

    row = table.add_row()
    set_cell_border(row.cells[0]); set_cell_bg(row.cells[0], 'D9D9D9'); add_cell_text(row.cells[0], 'الاجمالي')
    set_cell_border(row.cells[1]); add_cell_text(row.cells[1], inv.get('الاجمالي', ''))
    set_cell_border(row.cells[2]); set_cell_bg(row.cells[2], 'D9D9D9'); add_cell_text(row.cells[2], 'IT :')
    set_cell_border(row.cells[3]); add_cell_text(row.cells[3], '')

    row = table.add_row()
    cell = row.cells[0].merge(row.cells[3])
    set_cell_border(cell)
    add_cell_text(cell, NOTE, bold=True, size=9)

def parse_invoices(text):
    text = re.sub(r'^/invoice\s*', '', text.strip(), flags=re.IGNORECASE)
    blocks = re.split(r'\n\s*---+\s*\n', text)
    invoices = []
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        inv = {}
        for line in block.splitlines():
            line = line.strip()
            for field in FIELDS:
                if line.startswith(field + ':'):
                    inv[field] = line[len(field)+1:].strip()
                    break
        if inv:
            invoices.append(inv)
    return invoices

def create_word(invoices):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21); section.page_height = Cm(29.7)
    section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Cm(1)
    for i, inv in enumerate(invoices):
        make_invoice_table(doc, inv)
        if i < len(invoices) - 1:
            doc.add_paragraph()
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ====== هاندلرز ======

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id in authorized_users:
        if user_id in user_templates:
            await update.message.reply_text("أهلاً! 👋\nابعت /invoice لعمل فواتير\nأو /setup لتغيير شكل الفاتورة")
        else:
            await update.message.reply_text("أهلاً! 👋\nابعت /setup عشان تبعتلي شكل الفاتورة بتاعتك")
    else:
        await update.message.reply_text(
            "أهلاً! 👋\n\n"
            "🔐 البوت ده خاص، محتاج كود دخول\n\n"
            "ابعتلي الكود بتاعك:"
        )
        return WAITING_CODE

async def check_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    code = update.message.text.strip()
    if code in valid_codes:
        authorized_users[user_id] = code
        save_codes()
        await update.message.reply_text(
            f"✅ الكود صح! أهلاً بيك\n\n"
            f"ابعت /setup عشان تبعتلي شكل الفاتورة بتاعتك"
        )
        return ConversationHandler.END
    else:
        await update.message.reply_text("❌ الكود غلط! جرب تاني أو تواصل مع الأدمن")
        return WAITING_CODE

async def setup(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in authorized_users:
        await update.message.reply_text("❌ محتاج كود دخول الأول! ابعت /start")
        return
    await update.message.reply_text(
        "📄 ابعتلي ملف Word شكل الفاتورة بتاعتك 👇"
    )
    return WAITING_TEMPLATE

async def receive_template(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message.document or not update.message.document.file_name.endswith('.docx'):
        await update.message.reply_text("❌ ابعتلي ملف Word بامتداد .docx")
        return WAITING_TEMPLATE
    user_id = update.effective_user.id
    template_file = await context.bot.get_file(update.message.document.file_id)
    buf = io.BytesIO()
    await template_file.download_to_memory(buf)
    user_templates[user_id] = buf.getvalue()
    await update.message.reply_text("✅ تمام! استلمت شكل الفاتورة\n\nدلوقتي ابعت /invoice مع البيانات")
    return ConversationHandler.END

async def invoice_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text or ""
    if user_id not in authorized_users:
        await update.message.reply_text("❌ محتاج كود دخول! ابعت /start")
        return
    if user_id not in user_templates:
        await update.message.reply_text("❌ ابعت /setup الأول عشان تبعتلي شكل الفاتورة")
        return
    if len(text.strip()) < 12:
        await update.message.reply_text("ابعت البيانات مع الأمر /invoice")
        return
    await update.message.reply_text("⏳ بشتغل على الفواتير...")
    invoices = parse_invoices(text)
    if not invoices:
        await update.message.reply_text("❌ مش لاقي فواتير!")
        return
    try:
        buf = create_word(invoices)
        await update.message.reply_document(
            document=buf,
            filename=f"فواتير_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            caption=f"✅ تم إنشاء *{len(invoices)} فاتورة* بنجاح! 🐉",
            parse_mode='Markdown'
        )
    except Exception as e:
        await update.message.reply_text(f"❌ خطأ: {e}")

# ====== أوامر الأدمن ======
async def addcode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ مش أدمن!")
        return
    if not context.args:
        await update.message.reply_text("استخدام: /addcode الكود")
        return
    code = context.args[0]
    valid_codes[code] = datetime.now().strftime('%Y-%m-%d')
    save_codes()
    await update.message.reply_text(f"✅ تم إضافة الكود: `{code}`", parse_mode='Markdown')

async def removecode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ مش أدمن!")
        return
    if not context.args:
        await update.message.reply_text("استخدام: /removecode الكود")
        return
    code = context.args[0]
    if code in valid_codes:
        del valid_codes[code]
        # احذف اليوزرز اللي عندهم الكود ده
        to_remove = [uid for uid, c in authorized_users.items() if c == code]
        for uid in to_remove:
            del authorized_users[uid]
        save_codes()
        await update.message.reply_text(f"✅ تم حذف الكود: `{code}`", parse_mode='Markdown')
    else:
        await update.message.reply_text("❌ الكود مش موجود")

async def listcodes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ مش أدمن!")
        return
    if not valid_codes:
        await update.message.reply_text("مفيش أكواد حالياً")
        return
    text = "📋 *الأكواد الشغالة:*\n\n"
    for code, date in valid_codes.items():
        users = [uid for uid, c in authorized_users.items() if c == code]
        text += f"• `{code}` - أضيف {date} - {len(users)} مستخدم\n"
    await update.message.reply_text(text, parse_mode='Markdown')

async def listusers(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ مش أدمن!")
        return
    if not authorized_users:
        await update.message.reply_text("مفيش مستخدمين حالياً")
        return
    text = f"👥 *المستخدمين: {len(authorized_users)}*\n\n"
    for uid, code in authorized_users.items():
        text += f"• ID: `{uid}` - كود: `{code}`\n"
    await update.message.reply_text(text, parse_mode='Markdown')

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in authorized_users:
        await update.message.reply_text("ابعت /start للبدء")
    else:
        await update.message.reply_text("ابعت /invoice لعمل فواتير أو /setup لتغيير الفاتورة")

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    code_conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={WAITING_CODE: [MessageHandler(filters.TEXT & ~filters.COMMAND, check_code)]},
        fallbacks=[]
    )

    template_conv = ConversationHandler(
        entry_points=[CommandHandler("setup", setup)],
        states={WAITING_TEMPLATE: [MessageHandler(filters.Document.ALL, receive_template)]},
        fallbacks=[]
    )

    app.add_handler(code_conv)
    app.add_handler(template_conv)
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
