#!/usr/bin/env python3
import os, re, logging
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

BOT_TOKEN = "8856222347:AAFLpWdMsW69wGgnWLsuvChF3RaV6hZlHeY"

logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)

FIELDS = ['الاسم', 'الرقم', 'العنوان', 'اللون', 'العرض', 'المقاس', 'العدد', 'الاجمالي']

NOTE = 'في حالة عدم استلام الاوردر يتم دفع مصريف الشحن كامله للمندوب.'

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

def add_cell_text(cell, text, bold=True, size=10, rtl=True):
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = para.add_run(str(text or ''))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = 'Arial'
    if rtl:
        pPr = para._p.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        pPr.append(bidi)

def make_invoice_table(doc, inv):
    table = doc.add_table(rows=0, cols=4)
    table.style = 'Table Grid'

    # تحديد عرض الأعمدة
    for col in table.columns:
        for cell in col.cells:
            cell.width = Cm(4)

    def add_row(data):
        row = table.add_row()
        for i, (text, gray, span) in enumerate(data):
            if i >= len(row.cells):
                break
            cell = row.cells[i]
            set_cell_border(cell)
            if gray:
                set_cell_bg(cell, 'D9D9D9')
            add_cell_text(cell, text)
        return row

    # هيدر Dragon.s
    row = table.add_row()
    cell = row.cells[0]
    cell.merge(row.cells[3])
    set_cell_border(cell)
    set_cell_bg(cell, 'D9D9D9')
    add_cell_text(cell, 'Dragon.s', bold=True, size=12)

    # التاريخ
    row = table.add_row()
    cell = row.cells[0]
    cell.merge(row.cells[3])
    set_cell_border(cell)
    add_cell_text(cell, 'التاريخ/ رقم الاوردر/', bold=True)

    # الاسم
    row = table.add_row()
    set_cell_border(row.cells[0])
    set_cell_bg(row.cells[0], 'D9D9D9')
    add_cell_text(row.cells[0], 'الاسم:')
    cell = row.cells[1]
    cell.merge(row.cells[3])
    set_cell_border(cell)
    add_cell_text(cell, inv.get('الاسم', ''))

    # الرقم
    row = table.add_row()
    set_cell_border(row.cells[0])
    set_cell_bg(row.cells[0], 'D9D9D9')
    add_cell_text(row.cells[0], 'الرقم')
    cell = row.cells[1]
    cell.merge(row.cells[3])
    set_cell_border(cell)
    add_cell_text(cell, inv.get('الرقم', ''))

    # العنوان
    row = table.add_row()
    set_cell_border(row.cells[0])
    set_cell_bg(row.cells[0], 'D9D9D9')
    add_cell_text(row.cells[0], 'العنوان')
    cell = row.cells[1]
    cell.merge(row.cells[3])
    set_cell_border(cell)
    add_cell_text(cell, inv.get('العنوان', ''))

    # اللون والعرض
    row = table.add_row()
    set_cell_border(row.cells[0])
    set_cell_bg(row.cells[0], 'D9D9D9')
    add_cell_text(row.cells[0], 'اللون:')
    set_cell_border(row.cells[1])
    add_cell_text(row.cells[1], inv.get('اللون', ''))
    set_cell_border(row.cells[2])
    set_cell_bg(row.cells[2], 'D9D9D9')
    add_cell_text(row.cells[2], 'العرض:')
    set_cell_border(row.cells[3])
    add_cell_text(row.cells[3], inv.get('العرض', ''))

    # المقاس والعدد
    row = table.add_row()
    set_cell_border(row.cells[0])
    set_cell_bg(row.cells[0], 'D9D9D9')
    add_cell_text(row.cells[0], 'المقاس')
    set_cell_border(row.cells[1])
    add_cell_text(row.cells[1], inv.get('المقاس', ''))
    set_cell_border(row.cells[2])
    set_cell_bg(row.cells[2], 'D9D9D9')
    add_cell_text(row.cells[2], 'العدد:')
    set_cell_border(row.cells[3])
    add_cell_text(row.cells[3], inv.get('العدد', ''))

    # الاجمالي
    row = table.add_row()
    set_cell_border(row.cells[0])
    set_cell_bg(row.cells[0], 'D9D9D9')
    add_cell_text(row.cells[0], 'الاجمالي')
    set_cell_border(row.cells[1])
    add_cell_text(row.cells[1], inv.get('الاجمالي', ''))
    set_cell_border(row.cells[2])
    set_cell_bg(row.cells[2], 'D9D9D9')
    add_cell_text(row.cells[2], 'IT :')
    set_cell_border(row.cells[3])
    add_cell_text(row.cells[3], '')

    # ملاحظة
    row = table.add_row()
    cell = row.cells[0]
    cell.merge(row.cells[3])
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
    # إعداد الصفحة
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1)
    section.right_margin = Cm(1)
    section.top_margin = Cm(1)
    section.bottom_margin = Cm(1)

    for i, inv in enumerate(invoices):
        make_invoice_table(doc, inv)
        if i < len(invoices) - 1:
            doc.add_paragraph()

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

HELP = """
🐉 *بوت Dragon.s للفواتير*

ابعت /invoice وبعدين البيانات:

```
/invoice
الاسم: عمر حجاج
الرقم: 01008395565
العنوان: كفر الشيخ دسوق
اللون: ازرق
العرض: ترنج مشجر
المقاس: لارج
العدد: 1
الاجمالي: 300
---
الاسم: سماح محمد
الرقم: 01017572671
العنوان: شارع فيصل المريوطيه
اللون: 2 اسود
العرض: ترنج مشجر
المقاس: مقاس 3 سنين
العدد: 2
الاجمالي: 490
```
"""

async def start(update, context):
    await update.message.reply_text("أهلاً! 👋\nابعت /invoice لعمل فواتير")

async def help_cmd(update, context):
    await update.message.reply_text(HELP, parse_mode='Markdown')

async def invoice_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text or ""
    if len(text.strip()) < 12:
        await update.message.reply_text(HELP, parse_mode='Markdown')
        return

    await update.message.reply_text("⏳ بشتغل على الفواتير...")

    invoices = parse_invoices(text)
    if not invoices:
        await update.message.reply_text("❌ مش لاقي فواتير!")
        return

    try:
        buf = create_word(invoices)
        filename = f"Dragon.s_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
        await update.message.reply_document(
            document=buf,
            filename=filename,
            caption=f"✅ تم إنشاء *{len(invoices)} فاتورة* بنجاح! 🐉",
            parse_mode='Markdown'
        )
    except Exception as e:
        await update.message.reply_text(f"❌ خطأ: {e}")

async def handle_message(update, context):
    await update.message.reply_text("ابعت /invoice لعمل فواتير 🙂")

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("invoice", invoice_cmd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("✅ البوت شغال!")
    app.run_polling()

if __name__ == '__main__':
    main()
