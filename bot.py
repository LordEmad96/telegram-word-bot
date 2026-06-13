import telebot
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import io
import os
import re
import time
import hashlib
from datetime import datetime
import numpy as np
import cv2
from pdf2image import convert_from_bytes
import zipfile
import threading
from collections import defaultdict
import logging
import gc

# ========== تنظیمات ==========
logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

TOKEN = os.environ.get("TOKEN")
if not TOKEN:
    raise ValueError("توکن تنظیم نشده است!")

bot = telebot.TeleBot(TOKEN)

MAX_FILE_SIZE = 10 * 1024 * 1024
SUPPORTED_IMAGE_FORMATS = ('.png', '.jpg', '.jpeg', '.bmp', '.tiff')
OCR_LANGUAGES = 'fas+eng'

stats = defaultdict(int)
stats_lock = threading.Lock()

# ========== پردازش تصویر ==========
class ImageProcessor:
    @staticmethod
    def preprocess_for_ocr(image):
        try:
            if isinstance(image, Image.Image):
                img_array = np.array(image.convert('RGB'))
            else:
                img_array = image
            
            h, w = img_array.shape[:2]
            if h > 1000 or w > 1000:
                scale = 1000 / max(h, w)
                new_w = int(w * scale)
                new_h = int(h * scale)
                img_array = cv2.resize(img_array, (new_w, new_h))
            
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            gray = cv2.medianBlur(gray, 3)
            _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
            
            return Image.fromarray(binary)
        except Exception as e:
            logger.error(f"خطا در پیش‌پردازش: {e}")
            return image

# ========== پردازش اسناد ==========
class DocumentProcessor:
    @staticmethod
    def process_image(image_bytes):
        try:
            img = Image.open(io.BytesIO(image_bytes))
            if img.size[0] > 1200 or img.size[1] > 1200:
                img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
            
            processed = ImageProcessor.preprocess_for_ocr(img)
            text = pytesseract.image_to_string(processed, lang=OCR_LANGUAGES, config='--psm 6')
            
            text = re.sub(r'\s+', ' ', text).strip()
            
            del img, processed
            gc.collect()
            
            return text, len(text) > 50
        except Exception as e:
            logger.error(f"خطا در پردازش تصویر: {e}")
            return "", False
    
    @staticmethod
    def process_pdf(pdf_bytes, progress_callback=None):
        all_pages = []
        try:
            images = convert_from_bytes(pdf_bytes, dpi=150, fmt='jpeg')
            images = images[:10]
            
            for i, img in enumerate(images):
                if progress_callback:
                    progress_callback(i + 1, len(images))
                
                if img.size[0] > 1000 or img.size[1] > 1000:
                    img.thumbnail((1000, 1000), Image.Resampling.LANCZOS)
                
                img_byte = io.BytesIO()
                img.save(img_byte, format='JPEG', quality=60)
                
                text, success = DocumentProcessor.process_image(img_byte.getvalue())
                
                all_pages.append({
                    'page': i + 1,
                    'text': text[:2000] if success else '[متنی تشخیص داده نشد]',
                    'has_text': success
                })
                
                del img, img_byte
                gc.collect()
            
            return all_pages
        except Exception as e:
            logger.error(f"خطا در پردازش PDF: {e}")
            return []

# ========== هندلرهای ربات ==========
@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(message, """
🤖 **ربات تبدیل فایل به Word**

📤 ارسال کنید:
• عکس 📷
• PDF 📚
• ZIP 📦

⚠️ محدودیت: حداکثر 10 مگابایت
""", parse_mode='Markdown')

@bot.message_handler(commands=['stats'])
def show_stats(message):
    bot.reply_to(message, f"""
📊 **آمار**
• موفق: {stats.get('successful', 0)}
• ناموفق: {stats.get('failed', 0)}
• مجموع: {stats.get('total_processed', 0)}
""", parse_mode='Markdown')

@bot.message_handler(content_types=['photo'])
def handle_photo(message):
    chat_id = message.chat.id
    msg = bot.reply_to(message, "📷 در حال پردازش...")
    
    try:
        photo = message.photo[-1]
        file_info = bot.get_file(photo.file_id)
        downloaded = bot.download_file(file_info.file_path)
        
        text, success = DocumentProcessor.process_image(downloaded)
        
        if success:
            doc = Document()
            doc.add_heading('متن استخراج شده', 0)
            doc.add_paragraph(f"تاریخ: {datetime.now().strftime('%Y/%m/%d %H:%M')}")
            doc.add_paragraph()
            
            for line in text.split('\n')[:30]:
                if line.strip():
                    p = doc.add_paragraph()
                    run = p.add_run(line.strip()[:500])
                    run.font.size = Pt(10)
            
            output_path = f"output_{chat_id}.docx"
            doc.save(output_path)
            
            with open(output_path, 'rb') as f:
                bot.send_document(chat_id, f, caption="✅ فایل Word")
            
            os.remove(output_path)
            stats['successful'] += 1
        else:
            bot.reply_to(message, "❌ متنی تشخیص داده نشد")
            stats['failed'] += 1
        
        stats['total_processed'] += 1
        bot.delete_message(chat_id, msg.message_id)
        
    except Exception as e:
        bot.reply_to(message, f"❌ خطا: {str(e)[:100]}")
        stats['failed'] += 1

@bot.message_handler(content_types=['document'])
def handle_document(message):
    chat_id = message.chat.id
    file_name = message.document.file_name.lower()
    file_size = message.document.file_size
    
    if file_size > MAX_FILE_SIZE:
        bot.reply_to(message, "❌ حجم فایل بیشتر از 10 مگابایت است")
        return
    
    if not (file_name.endswith('.pdf') or file_name.endswith('.zip') or 
            file_name.endswith(('.png', '.jpg', '.jpeg'))):
        bot.reply_to(message, "❌ فرمت پشتیبانی نمی‌شود")
        return
    
    msg = bot.reply_to(message, f"📄 در حال پردازش {file_name}...")
    
    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded = bot.download_file(file_info.file_path)
        
        doc = Document()
        doc.add_heading('متن استخراج شده', 0)
        doc.add_paragraph(f"فایل: {file_name}")
        doc.add_paragraph(f"تاریخ: {datetime.now().strftime('%Y/%m/%d %H:%M')}")
        doc.add_paragraph()
        
        if file_name.endswith('.pdf'):
            def progress(current, total):
                bot.edit_message_text(f"📄 پردازش صفحه {current} از {total}", 
                                     chat_id, msg.message_id)
            
            pages = DocumentProcessor.process_pdf(downloaded, progress)
            for page in pages:
                if page['has_text']:
                    doc.add_heading(f"صفحه {page['page']}", level=2)
                    doc.add_paragraph(page['text'])
            stats['successful'] += 1
            
        elif file_name.endswith('.zip'):
            bot.edit_message_text("📦 در حال استخراج ZIP...", chat_id, msg.message_id)
            with zipfile.ZipFile(io.BytesIO(downloaded)) as z:
                for i, filename in enumerate(z.namelist()[:5]):
                    if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                        with z.open(filename) as f:
                            text, _ = DocumentProcessor.process_image(f.read())
                            if text:
                                doc.add_heading(f"فایل: {filename}", level=2)
                                doc.add_paragraph(text[:1000])
            stats['successful'] += 1
            
        else:
            text, success = DocumentProcessor.process_image(downloaded)
            if success:
                doc.add_paragraph(text)
                stats['successful'] += 1
            else:
                doc.add_paragraph("❌ متنی تشخیص داده نشد")
                stats['failed'] += 1
        
        output_path = f"output_{chat_id}.docx"
        doc.save(output_path)
        
        with open(output_path, 'rb') as f:
            bot.send_document(chat_id, f, caption="✅ فایل Word آماده است")
        
        os.remove(output_path)
        stats['total_processed'] += 1
        bot.delete_message(chat_id, msg.message_id)
        
    except Exception as e:
        bot.reply_to(message, f"❌ خطا: {str(e)[:100]}")
        stats['failed'] += 1

@bot.message_handler(func=lambda m: True)
def unknown(message):
    bot.reply_to(message, "عکس یا PDF بفرستید. /start برای راهنما")

# ========== راه‌اندازی ==========
if __name__ == '__main__':
    print("=" * 40)
    print("ربات راه‌اندازی شد!")
    print("=" * 40)
    bot.infinity_polling(timeout=60)
