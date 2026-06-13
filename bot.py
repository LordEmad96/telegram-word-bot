import telebot
from docx import Document
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import io
import os
from pdf2image import convert_from_bytes
import zipfile
import numpy as np

TOKEN = os.environ.get("TOKEN")
bot = telebot.TeleBot(TOKEN)

def enhance_image(img):
    """بهبود کیفیت تصویر برای OCR بهتر"""
    try:
        # تبدیل به RGB
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        # افزایش کنتراست
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2.0)
        
        # افزایش شارپنس (وضوح)
        enhancer = ImageEnhance.Sharpness(img)
        img = enhancer.enhance(2.0)
        
        # تبدیل به سیاه و سفید با آستانه‌گذاری
        gray = img.convert('L')
        
        # اعمال فیلتر میانه برای کاهش نویز
        gray = gray.filter(ImageFilter.MedianFilter(size=3))
        
        # آستانه‌گذاری (binary thresholding)
        threshold = 150
        img_binary = gray.point(lambda p: 255 if p > threshold else 0, '1')
        
        # بزرگنمایی خفیف
        width, height = img_binary.size
        img_binary = img_binary.resize((int(width*1.2), int(height*1.2)), Image.Resampling.LANCZOS)
        
        return img_binary
    except Exception as e:
        print(f"خطا در بهبود تصویر: {e}")
        return img

def extract_text_from_image(img):
    """استخراج متن از تصویر با تنظیمات بهینه"""
    try:
        # بهبود تصویر
        img_enhanced = enhance_image(img)
        
        # تنظیمات پیشرفته Tesseract
        custom_config = r'--oem 3 --psm 6 -l fas+eng'
        text = pytesseract.image_to_string(img_enhanced, config=custom_config)
        
        # حذف فاصله‌های اضافی و خطوط خالی
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        return '\n'.join(lines)
    except Exception as e:
        print(f"خطا در OCR: {e}")
        return ""

def process_pdf_to_text(pdf_bytes):
    """پردازش PDF و استخراج متن از هر صفحه"""
    all_text = []
    try:
        # تبدیل PDF به تصاویر با کیفیت بالا
        images = convert_from_bytes(pdf_bytes, dpi=300)  # افزایش DPI به 300
        
        for i, img in enumerate(images):
            page_text = extract_text_from_image(img)
            if page_text.strip():
                all_text.append(f"===== صفحه {i+1} =====\n{page_text}")
            else:
                all_text.append(f"===== صفحه {i+1} =====\n[متنی تشخیص داده نشد]")
    
    except Exception as e:
        all_text.append(f"خطا در پردازش PDF: {e}")
    
    return all_text

@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(message, 
        "سلام! من ربات تبدیل فایل به Word هستم.\n\n"
        "📤 می‌توانی این فایل‌ها را بفرستی:\n"
        "• عکس (JPG, PNG)\n"
        "• PDF (اسکن شده یا متنی)\n"
        "• فایل‌های زیپ شامل عکس\n\n"
        "🔍 برای بهترین نتیجه:\n"
        "• عکس با نور کافی و صاف گرفته شود\n"
        "• کیفیت تصویر بالا باشد\n"
        "• فونت متن ساده و خوانا باشد")

@bot.message_handler(content_types=['photo', 'document'])
def handle_files(message):
    chat_id = message.chat.id
    bot.send_message(chat_id, "⚙️ در حال پردازش... این ممکن است چند لحظه طول بکشد.")
    
    doc = Document()
    doc.add_heading('متن استخراج شده از فایل', 0)
    
    text_content = []
    
    try:
        # پردازش عکس مستقیم
        if message.photo:
            file_info = bot.get_file(message.photo[-1].file_id)
            downloaded = bot.download_file(file_info.file_path)
            img = Image.open(io.BytesIO(downloaded))
            text = extract_text_from_image(img)
            if text:
                text_content.append(text)
            else:
                text_content.append("[متنی در این عکس یافت نشد]")
        
        # پردازش فایل ضمیمه شده
        elif message.document:
            file_info = bot.get_file(message.document.file_id)
            downloaded = bot.download_file(file_info.file_path)
            file_name = message.document.file_name.lower()
            
            if file_name.endswith('.pdf'):
                bot.send_message(chat_id, "📄 در حال پردازش PDF... (ممکن است 1-2 دقیقه طول بکشد)")
                text_content = process_pdf_to_text(downloaded)
                
            elif file_name.endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
                img = Image.open(io.BytesIO(downloaded))
                text = extract_text_from_image(img)
                if text:
                    text_content.append(text)
                else:
                    text_content.append("[متنی در این عکس یافت نشد]")
                
            elif file_name.endswith('.zip'):
                bot.send_message(chat_id, "📦 در حال استخراج فایل‌های زیپ...")
                with zipfile.ZipFile(io.BytesIO(downloaded)) as z:
                    for file_in_zip in z.namelist():
                        if file_in_zip.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.pdf')):
                            bot.send_message(chat_id, f"🔄 پردازش: {file_in_zip}")
                            with z.open(file_in_zip) as f:
                                if file_in_zip.lower().endswith('.pdf'):
                                    pdf_bytes = f.read()
                                    pdf_text = process_pdf_to_text(pdf_bytes)
                                    text_content.append(f"\n--- فایل: {file_in_zip} ---")
                                    text_content.extend(pdf_text)
                                else:
                                    img = Image.open(f)
                                    text = extract_text_from_image(img)
                                    if text:
                                        text_content.append(f"\n--- فایل: {file_in_zip} ---\n{text}")
            else:
                bot.send_message(chat_id, "❌ فرمت فایل پشتیبانی نمی‌شود. فقط عکس، PDF و زیپ.")
                return
        
        # ساختن فایل Word
        if text_content:
            for paragraph in text_content:
                if paragraph.strip():
                    doc.add_paragraph(paragraph)
                    doc.add_paragraph()  # فاصله بین پاراگراف‌ها
        else:
            doc.add_paragraph("متنی استخراج نشد. لطفاً فایل با کیفیت بهتری ارسال کنید.")
        
        # ذخیره و ارسال فایل
        output_path = f"output_{chat_id}.docx"
        doc.save(output_path)
        
        with open(output_path, 'rb') as f:
            bot.send_document(chat_id, f, caption="✅ فایل Word با موفقیت ساخته شد!\n\n📌 توجه: کیفیت خروجی به کیفیت فایل اصلی بستگی دارد.")
        
        os.remove(output_path)
        bot.send_message(chat_id, "🎉 انجام شد! اگر راضی بودی می‌توانی ربات را به دوستانت معرفی کنی.")
        
    except Exception as e:
        bot.send_message(chat_id, f"❌ خطایی رخ داد: {str(e)[:100]}\n\nلطفاً فایل با کیفیت بهتری ارسال کنید یا دوباره تلاش کنید.")
        print(f"خطا: {e}")

# راه‌اندازی ربات
print("🤖 ربات روشن شد...")
print("📱 منتظر دریافت فایل‌ها...")
bot.infinity_polling()
