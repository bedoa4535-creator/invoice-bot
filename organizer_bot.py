#!/usr/bin/env python3
import os, logging, re
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler
import anthropic
import io

BOT_TOKEN = os.environ.get("BOT_TOKEN2")
CLAUDE_API_KEY = os.environ.get("CLAUDE_API_KEY")

logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)

WAITING_TEMPLATE = 1
user_templates = {}

TEMPLATES = {
    'dragon': ['الاسم', 'الرقم', 'العنوان', 'اللون', 'العرض', 'المقاس', 'العدد', 'الاجمالي'],
    'fashion': ['اسم العميل', 'اكونت العميل', 'رقم الفاتورة', 'اسم الصفحة', 'اسم IT', 'العنوان', 'الرقم', 'الألوان', 'العدد', 'المقاس', 'نوع المنتج', 'سعر المنتج', 'مصاريف الشحن', 'الاجمالي'],
    'mm': ['اسم الاكونت', 'IT', 'الاسم', 'العنوان', 'موبايل 1', 'موبايل 2', 'المنتج', 'الالوان', 'المقاس', 'الكمية', 'السعر', 'الشحن', 'الاجمالي'],
}

def organize_with_claude(raw_text, template_type):
    fields = TEMPLATES[template_type]
    fields_str = '\n'.join([f'- {f}' for f in fields])

    client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)

    prompt = f"""أنت مساعد متخصص في تنظيم بيانات الفواتير.

البيانات الواردة قد تكون مكتوبة بأي شكل عشوائي وبها رموز مختلفة مثل / أو // أو - أو مسافات.

مهمتك:
1. اقرأ البيانات الخام
2. استخرج المعلومات المطلوبة
3. رتبها بالشكل الصح

الحقول المطلوبة:
{fields_str}

قواعد مهمة:
- كل فاتورة تنتهي بـ ---
- لو مفيش قيمة لحقل معين اكتب: الحقل: غير محدد
- الرقم: اكتب الأرقام بدون مسافات
- لا تضيف أي كلام غير البيانات المرتبة
- افصل كل فاتورة عن التانية بـ ---

البيانات الخام:
{raw_text}

اكتب البيانات المرتبة فقط بدون أي مقدمة أو كلام إضافي:"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )

    return message.content[0].text

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "أهلاً! 👋 أنا بوت تنظيم الفواتير\n\n"
        "بتبعتلي بيانات الفواتير بأي شكل وأنا برتبهالك\n\n"
        "أول حاجة، قولي هتبعت فواتير لأنهي شكل؟\n\n"
        "• اكتب *dragon*\n"
        "• اكتب *fashion*\n"
        "• اكتب *mm*",
        parse_mode='Markdown'
    )
    return WAITING_TEMPLATE

async def choose_template(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    choice = update.message.text.strip().lower()
    if choice in TEMPLATES:
        user_templates[user_id] = choice
        await update.message.reply_text(
            f"✅ تمام! هترتب الفواتير لشكل *{choice}*\n\n"
            "دلوقتي ابعتلي البيانات كلام عادي أو ملف .txt\n"
            "وأنا هرتبهالك فوراً 🚀",
            parse_mode='Markdown'
        )
        return ConversationHandler.END
    else:
        await update.message.reply_text(
            "❌ مش عارف الشكل ده!\n\n"
            "اكتب:\n• *dragon*\n• *fashion*\n• *mm*",
            parse_mode='Markdown'
        )
        return WAITING_TEMPLATE

async def process_text(raw_text, update, template_type):
    await update.message.reply_text("⏳ بنظم الفواتير...")
    try:
        organized = organize_with_claude(raw_text, template_type)
        # لو النص قصير نبعته كرسالة
        if len(organized) < 3000:
            await update.message.reply_text(
                f"✅ *الفواتير المنظمة:*\n\n```\n{organized}\n```",
                parse_mode='Markdown'
            )
        # لو النص طويل نبعته كملف txt
        buf = io.BytesIO(organized.encode('utf-8'))
        buf.name = f"فواتير_منظمة_{template_type}.txt"
        await update.message.reply_document(
            document=buf,
            filename=f"فواتير_منظمة_{template_type}.txt",
            caption="✅ الفواتير المنظمة جاهزة! ابعتها للبوت التاني 🎉"
        )
    except Exception as e:
        await update.message.reply_text(f"❌ خطأ: {e}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    template_type = user_templates.get(user_id)
    if not template_type:
        await update.message.reply_text("ابعت /start الأول عشان تختار شكل الفاتورة")
        return
    raw_text = update.message.text
    await process_text(raw_text, update, template_type)

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    template_type = user_templates.get(user_id)
    if not template_type:
        await update.message.reply_text("ابعت /start الأول عشان تختار شكل الفاتورة")
        return
    doc = update.message.document
    if not doc.file_name.endswith('.txt'):
        await update.message.reply_text("❌ ابعتلي ملف .txt بس")
        return
    file = await context.bot.get_file(doc.file_id)
    buf = io.BytesIO()
    await file.download_to_memory(buf)
    try:
        raw_text = buf.getvalue().decode('utf-8')
    except:
        raw_text = buf.getvalue().decode('windows-1256', errors='ignore')
    await process_text(raw_text, update, template_type)

async def change_template(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "اختار شكل الفاتورة:\n\n"
        "• *dragon*\n• *fashion*\n• *mm*",
        parse_mode='Markdown'
    )
    return WAITING_TEMPLATE

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            CommandHandler("change", change_template),
        ],
        states={
            WAITING_TEMPLATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, choose_template)],
        },
        fallbacks=[]
    )
    app.add_handler(conv)
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("✅ بوت التنظيم شغال!")
    app.run_polling()

if __name__ == '__main__':
    main()
