import logging
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from docx import Document
import os
from docx2pdf import convert
from PIL import Image

BOT_TOKEN = "ضع_التوكن_هنا"
TEMPLATE_FILE = "template.docx"  # ملف الشهادة الأساسي

# 🔹 إعداد اللوجز
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

# 🟢 تحويل الرقم إلى كلمة ترتيبية بالعربية
def arabic_order(n):
    orders = {
        1: "الأول",
        2: "الثاني",
        3: "الثالث",
        4: "الرابع",
        5: "الخامس",
        6: "السادس",
        7: "السابع",
        8: "الثامن",
        9: "التاسع",
        10: "العاشر",
    }
    return orders.get(n, f"الرقم {n}")

# 🟢 /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 أهلاً بك!\n"
        "أرسل لي قائمة أسماء (كل اسم في سطر جديد)\n"
        "وسأقوم بإنشاء شهادة شكر مستقلة لكل شخص مع الترتيب 🎓."
    )

# 🟢 إنشاء شهادة من ملف Word
def create_certificate(name, order, index):
    # فتح القالب
    doc = Document(TEMPLATE_FILE)

    # استبدال النصوص (ضع في القالب كلمة مفتاحية مثل {{NAME}} و {{ORDER}})
    for p in doc.paragraphs:
        if "{{NAME}}" in p.text:
            p.text = p.text.replace("{{NAME}}", name)
        if "{{ORDER}}" in p.text:
            p.text = p.text.replace("{{ORDER}}", order)

    # حفظ كملف مؤقت
    out_docx = f"certificate_{index}.docx"
    out_pdf = f"certificate_{index}.pdf"
    out_img = f"certificate_{index}.png"

    doc.save(out_docx)

    # تحويل PDF
    convert(out_docx, out_pdf)

    # تحويل PDF إلى صورة (باستخدام Pillow + poppler أو pdf2image)
    from pdf2image import convert_from_path
    images = convert_from_path(out_pdf)
    images[0].save(out_img, "PNG")

    # تنظيف الملفات غير الضرورية
    os.remove(out_docx)
    os.remove(out_pdf)

    return out_img

# 🟢 استقبال الأسماء
async def names_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    names = text.split("\n")

    for i, name in enumerate(names, start=1):
        order_word = arabic_order(i)

        img_file = create_certificate(name, order_word, i)

        await update.message.reply_photo(photo=open(img_file, "rb"))
        os.remove(img_file)

# 🟢 تشغيل البوت
def main():
    application = Application.builder().token(BOT_TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, names_handler))

    application.run_polling()

if __name__ == "__main__":
    main()
