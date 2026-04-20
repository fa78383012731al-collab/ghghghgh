import os
import threading
import logging
from io import BytesIO

from flask import Flask
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, CallbackQueryHandler, CallbackContext
from pptx import Presentation

# ---------- الإعدادات ----------
BOT_TOKEN = os.getenv("BOT_TOKEN")
PORT = int(os.getenv("PORT", "10000"))

# إعداد التسجيل
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# قاموس مؤقت لحفظ الملفات
user_files = {}

# ---------- Flask ----------
app = Flask(__name__)

@app.route("/")
def home():
    return "✂️ PowerPoint Crop Bot is running."

# ---------- دوال معالجة PPTX ----------
def crop_pptx_from_bottom(file_bytes: bytes, crop_percent: int) -> BytesIO:
    prs = Presentation(BytesIO(file_bytes))
    original_width = prs.slide_width
    original_height = prs.slide_height
    new_height = int(original_height * (1 - crop_percent / 100.0))
    prs.slide_width = original_width
    prs.slide_height = new_height
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# ---------- أوامر البوت ----------
def start(update: Update, context: CallbackContext):
    update.message.reply_text(
        "🎬 أهلاً بك في بوت قص شرائح البوربوينت!\n\n"
        "📤 أرسل لي ملف PPTX لتبدأ.\n"
        "✂️ بعدها ستختار نسبة القص من الأسفل (1% - 80%)."
    )

def handle_document(update: Update, context: CallbackContext):
    user_id = update.effective_user.id
    document = update.message.document

    if not document.file_name.lower().endswith(".pptx"):
        update.message.reply_text("❌ الملف يجب أن يكون بصيغة .pptx فقط.")
        return

    file = context.bot.get_file(document.file_id)
    file_bytes = file.download_as_bytearray()
    user_files[user_id] = bytes(file_bytes)

    keyboard = [
        [
            InlineKeyboardButton("10%", callback_data="crop_10"),
            InlineKeyboardButton("20%", callback_data="crop_20"),
            InlineKeyboardButton("30%", callback_data="crop_30"),
        ],
        [
            InlineKeyboardButton("40%", callback_data="crop_40"),
            InlineKeyboardButton("50%", callback_data="crop_50"),
            InlineKeyboardButton("60%", callback_data="crop_60"),
        ],
        [
            InlineKeyboardButton("70%", callback_data="crop_70"),
            InlineKeyboardButton("80%", callback_data="crop_80"),
        ],
        [InlineKeyboardButton("✏️ إدخال نسبة يدوية", callback_data="manual_crop")],
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)

    update.message.reply_text(
        f"✅ تم استلام الملف: `{document.file_name}`\n\n"
        "🔽 اختر نسبة القص من الأسفل:",
        reply_markup=reply_markup,
        parse_mode="Markdown"
    )

def button_callback(update: Update, context: CallbackContext):
    query = update.callback_query
    query.answer()
    user_id = update.effective_user.id
    data = query.data

    if user_id not in user_files:
        query.edit_message_text("⚠️ لم يتم العثور على ملف. أرسل ملف PPTX أولاً.")
        return

    if data == "manual_crop":
        query.edit_message_text(
            "📝 الرجاء إرسال النسبة المطلوبة (رقم بين 1 و 80) في رسالة نصية:"
        )
        context.user_data["awaiting_crop_value"] = True
        return

    percent = int(data.split("_")[1])
    process_crop(update, context, user_id, percent, is_manual=False)

def handle_text(update: Update, context: CallbackContext):
    user_id = update.effective_user.id

    if not context.user_data.get("awaiting_crop_value"):
        return

    text = update.message.text.strip()
    try:
        percent = int(text)
        if percent < 1 or percent > 80:
            update.message.reply_text("❌ النسبة يجب أن تكون بين 1 و 80. حاول مجدداً:")
            return
    except ValueError:
        update.message.reply_text("❌ الرجاء إرسال رقم صحيح بين 1 و 80:")
        return

    context.user_data["awaiting_crop_value"] = False
    process_crop(update, context, user_id, percent, is_manual=True)

def process_crop(update: Update, context: CallbackContext, user_id: int, percent: int, is_manual: bool):
    file_bytes = user_files.get(user_id)
    if not file_bytes:
        if is_manual:
            update.message.reply_text("⚠️ انتهت صلاحية الملف. أرسل PPTX مجدداً.")
        else:
            update.callback_query.edit_message_text("⚠️ انتهت صلاحية الملف. أرسل PPTX مجدداً.")
        return

    if is_manual:
        msg = update.message.reply_text("⏳ جاري معالجة الملف...")
    else:
        msg = update.callback_query.edit_message_text("⏳ جاري معالجة الملف...")

    try:
        output_stream = crop_pptx_from_bottom(file_bytes, percent)

        context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=output_stream,
            filename=f"cropped_{percent}percent.pptx",
            caption=f"✅ تم قص {percent}% من أسفل الشرائح بنجاح!"
        )
        msg.delete()
    except Exception as e:
        logger.error(f"خطأ أثناء معالجة الملف: {e}")
        error_text = f"❌ حدث خطأ أثناء المعالجة: {str(e)}"
        if is_manual:
            msg.edit_text(error_text)
        else:
            msg.edit_text(error_text)
    finally:
        user_files.pop(user_id, None)

def error_handler(update: Update, context: CallbackContext):
    logger.error(msg="استثناء غير معالج:", exc_info=context.error)

# ---------- تشغيل البوت في خيط منفصل ----------
def run_bot():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN is missing")

    updater = Updater(BOT_TOKEN, use_context=True)
    dp = updater.dispatcher

    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(MessageHandler(Filters.document, handle_document))
    dp.add_handler(MessageHandler(Filters.text & ~Filters.command, handle_text))
    dp.add_handler(CallbackQueryHandler(button_callback))
    dp.add_error_handler(error_handler)

    logger.info("🤖 بوت قص البوربوينت يعمل الآن...")
    updater.start_polling()
    updater.idle()

if __name__ == "__main__":
    threading.Thread(target=run_bot, daemon=True).start()
    app.run(host="0.0.0.0", port=PORT)
