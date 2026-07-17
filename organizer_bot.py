#!/usr/bin/env python3
import os, logging, re, json
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler
import anthropic
import io

BOT_TOKEN = os.environ.get("BOT_TOKEN2")
CLAUDE_API_KEY = os.environ.get("CLAUDE_API_KEY")
ADMIN_ID = int(os.environ.get("ADMIN_ID", "0"))

logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)

WAITING_TEMPLATE = 1
user_templates = {}
STATS_FILE = "stats.json"

# ====== إحصائيات ======
def load_stats():
    if os.path.exists(STATS_FILE):
        try:
            with open(STATS_FILE) as f:
                return json.load(f)
        except:
            pass
    return {}

def save_stats(stats):
    with open(STATS_FILE, 'w') as f:
        json.dump(stats, f, ensure_ascii=False)

def update_stats(user_id, username, invoices_count):
    stats = load_stats()
    uid = str(user_id)
    if uid not in stats:
        stats[uid] = {
            'username': username or 'غير معروف',
            'total_invoices': 0,
            'requests': 0,
            'last_use': '',
            'first_use': datetime.now().strftime('%Y-%m-%d')
        }
    stats[uid]['total_invoices'] += invoices_count
    stats[uid]['requests'] += 1
    stats[uid]['last_use'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    stats[uid]['username'] = username or stats[uid]['username']
    save_stats(stats)

TEMPLATES = {
    'dragon': ['الاسم', 'الرقم', 'العنوان', 'اللون', 'العرض', 'المقاس', 'العدد', 'الاجمالي'],
    'fashion': ['اسم العميل', 'اكونت العميل', 'رقم الفاتورة', 'اسم الصفحة', 'اسم IT', 'العنوان', 'الرقم', 'الألوان', 'العدد', 'المقاس', 'نوع المنتج', 'سعر المنتج', 'مصاريف الشحن', 'الاجمالي'],
    'mm': ['اسم الاكونت', 'IT', 'الاسم', 'العنوان', 'موبايل 1', 'موبايل 2', 'المنتج', 'الالوان', 'المقاس', 'الكمية', 'السعر', 'الشحن', 'الاجمالي'],
}

def split_invoices(text, chunk_size=15):
    separators = [
        r'\[\d+/\d+,\s*\d+:\d+\s*[صم]\]',
        r'\+20\s*\d+',
        r'تاريخ حجز الاوردر',
        r'تاريخ حجز الاوردار',
        r'اسم العميل',
    ]
    lines = text.strip().split('\n')
    chunks = []
    current_chunk = []
    invoice_count = 0
    for line in lines:
        line = line.strip()
        if not line:
            continue
        is_new_invoice = any(re.search(sep, line) for sep in separators)
        if is_new_invoice and invoice_count > 0 and invoice_count % chunk_size == 0:
            chunks.append('\n'.join(current_chunk))
            current_chunk = []
        if is_new_invoice:
            invoice_count += 1
        current_chunk.append(line)
    if current_chunk:
        chunks.append('\n'.join(current_chunk))
    if len(chunks) <= 1 and len(text) > 3000:
        chunks = []
        words = text.split()
        chunk_words = len(words) // ((len(text) // 3000) + 1)
        for i in range(0, len(words), chunk_words):
            chunks.append(' '.join(words[i:i+chunk_words]))
    return chunks if chunks else [text]

def count_invoices(text):
    """عد الفواتير في النص المرتب"""
    return text.count('---')

def organize_chunk_with_claude(raw_text, template_type):
    fields = TEMPLATES[template_type]
    fields_str = '\n'.join([f'- {f}' for f in fields])
    client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)
    prompt = f"""أنت مساعد متخصص في تنظيم بيانات الفواتير.

البيانات الواردة قد تكون مكتوبة بأي شكل عشوائي.

مهمتك:
1. اقرأ البيانات الخام
2. استخرج كل فاتورة
3. رتب كل فاتورة بالشكل الصح

الحقول المطلوبة لكل فاتورة:
{fields_str}

قواعد مهمة:
- كل فاتورة تنتهي بـ ---
- لو مفيش قيمة لحقل اكتب: الحقل: غير محدد
- الرقم: اكتب الأرقام بدون مسافات
- لا تضيف أي كلام إضافي
- افصل كل فاتورة عن التانية بسطر فيه --- فقط

البيانات:
{raw_text}

اكتب الفواتير المرتبة فقط:"""
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )
    return message.content[0].text

async def process_invoices(raw_text, update, template_type):
    user_id = update.effective_user.id
    username = update.effective_user.username or update.effective_user.first_name

    await update.message.reply_text("⏳ بنظم الفواتير، استنى شوية...")
    try:
        chunks = split_invoices(raw_text)
        total_chunks = len(chunks)
        if total_chunks > 1:
            await update.message.reply_text(f"📦 الملف كبير، هشتغل عليه في {total_chunks} أجزاء...")
        all_organized = []
        for i, chunk in enumerate(chunks):
            if total_chunks > 1:
                await update.message.reply_text(f"⏳ بشتغل على الجزء {i+1} من {total_chunks}...")
            organized = organize_chunk_with_claude(chunk, template_type)
            all_organized.append(organized.strip())
        final_text = '\n---\n'.join(all_organized)

        # احسب عدد الفواتير
        invoices_count = count_invoices(final_text) + 1

        # حفظ الإحصائيات
        update_stats(user_id, username, invoices_count)

        buf = io.BytesIO(final_text.encode('utf-8'))
        await update.message.reply_document(
            document=buf,
            filename=f"فواتير_منظمة_{template_type}.txt",
            caption=f"✅ تم تنظيم *{invoices_count} فاتورة* بنجاح! جاهزة للبوت التاني 🎉",
            parse_mode='Markdown'
        )
        if len(final_text) < 3000:
            await update.message.reply_text(f"```\n{final_text}\n```", parse_mode='Markdown')
    except Exception as e:
        await update.message.reply_text(f"❌ خطأ: {e}")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "أهلاً! 👋 أنا بوت تنظيم الفواتير\n\n"
        "قولي هتبعت فواتير لأنهي شكل؟\n\n"
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
            f"✅ تمام! هرتب الفواتير لشكل *{choice}*\n\n"
            "ابعتلي البيانات كلام عادي أو ملف .txt 🚀",
            parse_mode='Markdown'
        )
        return ConversationHandler.END
    else:
        await update.message.reply_text(
            "❌ مش عارف الشكل ده!\n\nاكتب:\n• *dragon*\n• *fashion*\n• *mm*",
            parse_mode='Markdown'
        )
        return WAITING_TEMPLATE

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    template_type = user_templates.get(user_id)
    if not template_type:
        await update.message.reply_text("ابعت /start الأول عشان تختار شكل الفاتورة")
        return
    await process_invoices(update.message.text, update, template_type)

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
    await process_invoices(raw_text, update, template_type)

async def stats_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ مش أدمن!")
        return
    stats = load_stats()
    if not stats:
        await update.message.reply_text("مفيش إحصائيات لسه")
        return
    total_invoices = sum(v['total_invoices'] for v in stats.values())
    total_requests = sum(v['requests'] for v in stats.values())
    text = f"📊 *إحصائيات بوت الترتيب*\n\n"
    text += f"👥 إجمالي المستخدمين: *{len(stats)}*\n"
    text += f"🧾 إجمالي الفواتير: *{total_invoices}*\n"
    text += f"🔄 إجمالي الطلبات: *{total_requests}*\n\n"
    text += "━━━━━━━━━━━━━━\n"
    text += "*تفاصيل كل مستخدم:*\n\n"
    for uid, data in sorted(stats.items(), key=lambda x: x[1]['total_invoices'], reverse=True):
        text += f"👤 *{data['username']}*\n"
        text += f"   🧾 فواتير: {data['total_invoices']}\n"
        text += f"   🔄 طلبات: {data['requests']}\n"
        text += f"   📅 آخر استخدام: {data['last_use']}\n\n"
    await update.message.reply_text(text, parse_mode='Markdown')

async def mystats_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    stats = load_stats()
    uid = str(update.effective_user.id)
    if uid not in stats:
        await update.message.reply_text("مش عندك إحصائيات لسه، ابعت فواتير الأول!")
        return
    data = stats[uid]
    text = (
        f"📊 *إحصائياتك*\n\n"
        f"🧾 إجمالي الفواتير: *{data['total_invoices']}*\n"
        f"🔄 عدد الطلبات: *{data['requests']}*\n"
        f"📅 أول استخدام: {data['first_use']}\n"
        f"📅 آخر استخدام: {data['last_use']}"
    )
    await update.message.reply_text(text, parse_mode='Markdown')

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={WAITING_TEMPLATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, choose_template)]},
        fallbacks=[]
    )
    app.add_handler(conv)
    app.add_handler(CommandHandler("stats", stats_cmd))
    app.add_handler(CommandHandler("mystats", mystats_cmd))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("✅ بوت التنظيم شغال!")
    app.run_polling()

if __name__ == '__main__':
    main()
