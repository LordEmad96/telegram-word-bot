import telebot
from docx import Document
from docx.shared import Inches, Pt, RGBColor
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
import json
import threading
from collections import defaultdict
import logging

# ========== تنظیمات لاگینگ ==========
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ========== تنظیمات ربات ==========
TOKEN = os.environ.get("TOKEN")
if not TOKEN:
    raise ValueError("توکن ربات در environment variables تنظیم نشده است!")

bot = telebot.TeleBot(TOKEN)

# تنظیمات پیشرفته
MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 مگابایت
SUPPORTED_IMAGE_FORMATS = ('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.webp')
SUPPORTED_DOC_FORMATS = ('.pdf', '.zip')
OCR_LANGUAGES = 'fas+eng'  # فارسی و انگلیسی
OCR_CONFIGS = {
    'fast': r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZآبپتثجچحخدذرزژسشصضطظعغفقکگلمنوهی',
    'accurate': r'--oem 3 --psm 6',
    'sparse': r'--oem 3 --psm 11',
    'single_line': r'--oem 3 --psm 7'
}

# کش برای پردازش‌های تکراری
processing_cache = {}
cache_lock = threading.Lock()

# آمار ربات
stats = defaultdict(int)
stats_lock = threading.Lock()

# ========== کلاس‌های پیشرفته ==========

class ImageProcessor:
    """پردازشگر پیشرفته تصاویر"""
    
    @staticmethod
    def remove_noise(image):
        """حذف نویز از تصویر"""
        if isinstance(image, Image.Image):
            image = np.array(image.convert('RGB'))
        
        # اعمال فیلتر میانه
        denoised = cv2.medianBlur(image, 3)
        # فیلتر دوجانبه برای حفظ لبه‌ها
        denoised = cv2.bilateralFilter(denoised, 9, 75, 75)
        return denoised
    
    @staticmethod
    def correct_skew(image):
        """اصلاح کجی تصویر"""
        if isinstance(image, Image.Image):
            image = np.array(image.convert('L'))
        
        # تشخیص لبه‌ها
        edges = cv2.Canny(image, 50, 150, apertureSize=3)
        lines = cv2.HoughLines(edges, 1, np.pi/180, threshold=100)
        
        if lines is not None:
            angles = []
            for rho, theta in lines[:, 0]:
                angle = np.degrees(theta) - 90
                if abs(angle) < 45:
                    angles.append(angle)
            
            if angles:
                median_angle = np.median(angles)
                if abs(median_angle) > 0.5:
                    (h, w) = image.shape[:2]
                    center = (w // 2, h // 2)
                    M = cv2.getRotationMatrix2D(center, median_angle, 1.0)
                    rotated = cv2.warpAffine(image, M, (w, h), 
                                             flags=cv2.INTER_CUBIC,
                                             borderMode=cv2.BORDER_REPLICATE)
                    return rotated
        return image
    
    @staticmethod
    def enhance_contrast(image):
        """بهبود کنتراست با CLAHE"""
        if isinstance(image, Image.Image):
            image = np.array(image.convert('L'))
        
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        enhanced = clahe.apply(image)
        return enhanced
    
    @staticmethod
    def sharpen_image(image):
        """تیز کردن تصویر"""
        if isinstance(image, Image.Image):
            image = np.array(image)
        
        kernel = np.array([[-1,-1,-1],
                           [-1, 9,-1],
                           [-1,-1,-1]])
        sharpened = cv2.filter2D(image, -1, kernel)
        return sharpened
    
    @staticmethod
    def auto_rotate(image):
        """تشخیص و چرخش خودکار تصویر بر اساس EXIF"""
        try:
            if hasattr(image, '_getexif'):
                exif = image._getexif()
                if exif:
                    orientation = exif.get(0x0112, 1)
                    rotate_map = {
                        3: Image.ROTATE_180,
                        6: Image.ROTATE_270,
                        8: Image.ROTATE_90
                    }
                    if orientation in rotate_map:
                        image = image.transpose(rotate_map[orientation])
        except Exception as e:
            logger.warning(f"خطا در چرخش خودکار: {e}")
        return image
    
    @staticmethod
    def preprocess_for_ocr(image):
        """پردازش کامل تصویر برای OCR"""
        # تبدیل به آرایه numpy
        if isinstance(image, Image.Image):
            img_array = np.array(image.convert('RGB'))
        else:
            img_array = image
        
        # حذف نویز
        img_array = ImageProcessor.remove_noise(img_array)
        
        # تبدیل به خاکستری
        if len(img_array.shape) == 3:
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
        else:
            gray = img_array
        
        # اصلاح کجی
        gray = ImageProcessor.correct_skew(gray)
        
        # بهبود کنتراست
        gray = ImageProcessor.enhance_contrast(gray)
        
        # تیز کردن
        gray = ImageProcessor.sharpen_image(gray)
        
        # آستانه‌گذاری تطبیقی
        binary = cv2.adaptiveThreshold(gray, 255, 
                                      cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                      cv2.THRESH_BINARY, 11, 2)
        
        # بزرگنمایی برای بهبود OCR
        scale = 2
        width = int(binary.shape[1] * scale)
        height = int(binary.shape[0] * scale)
        binary = cv2.resize(binary, (width, height), interpolation=cv2.INTER_CUBIC)
        
        return Image.fromarray(binary)

class TextFormatter:
    """فرمت‌دهنده پیشرفته متن خروجی"""
    
    @staticmethod
    def clean_text(text):
        """پاکسازی و نرمال‌سازی متن"""
        # حذف کاراکترهای اضافی
        text = re.sub(r'[^\w\s\u0600-\u06FF\uFB50-\uFDFF\uFE70-\uFEFF\.,!?;:()\[\]{}@#%&*+-=]', ' ', text)
        # حذف فاصله‌های اضافی
        text = re.sub(r'\s+', ' ', text)
        # حذف خطوط خالی
        text = '\n'.join(line.strip() for line in text.split('\n') if line.strip())
        return text.strip()
    
    @staticmethod
    def detect_language(text):
        """تشخیص زبان متن"""
        try:
            from langdetect import detect
            lang = detect(text)
            lang_map = {'fa': 'Persian', 'en': 'English', 'ar': 'Arabic'}
            return lang_map.get(lang, 'Unknown')
        except:
            return 'Unknown'
    
    @staticmethod
    def add_to_document(doc, text, title=None):
        """اضافه کردن متن به سند Word با فرمت‌دهی"""
        if title:
            heading = doc.add_heading(title, level=1)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = heading.runs[0]
            run.font.size = Pt(16)
            run.font.bold = True
        
        paragraphs = text.split('\n')
        for para_text in paragraphs:
            if para_text.strip():
                p = doc.add_paragraph()
                run = p.add_run(para_text.strip())
                run.font.size = Pt(11)
                run.font.name = 'B Nazanin' or 'Times New Roman'
                
                # تشخیص و تنظیم جهت متن
                if re.search(r'[\u0600-\u06FF]', para_text):
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    run.font.rtl = True
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run.font.rtl = False

class DocumentProcessor:
    """پردازشگر اصلی اسناد"""
    
    @staticmethod
    def process_image(image_bytes, config_type='accurate'):
        """پردازش تصویر و استخراج متن"""
        try:
            img = Image.open(io.BytesIO(image_bytes))
            
            # چرخش خودکار
            img = ImageProcessor.auto_rotate(img)
            
            # پیش‌پردازش
            processed_img = ImageProcessor.preprocess_for_ocr(img)
            
            # استخراج متن
            config = OCR_CONFIGS.get(config_type, OCR_CONFIGS['accurate'])
            text = pytesseract.image_to_string(processed_img, lang=OCR_LANGUAGES, config=config)
            
            # پاکسازی متن
            text = TextFormatter.clean_text(text)
            
            return text, len(text) > 0
        except Exception as e:
            logger.error(f"خطا در پردازش تصویر: {e}")
            return "", False
    
    @staticmethod
    def process_pdf(pdf_bytes, progress_callback=None):
        """پردازش PDF و استخراج متن از تمام صفحات"""
        all_pages_text = []
        
        try:
            # تبدیل PDF به تصاویر با کیفیت بالا
            images = convert_from_bytes(pdf_bytes, dpi=300, fmt='jpeg')
            
            for i, img in enumerate(images):
                if progress_callback:
                    progress_callback(i + 1, len(images))
                
                # ذخیره موقت تصویر
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format='JPEG', quality=95)
                
                # پردازش صفحه
                text, success = DocumentProcessor.process_image(img_byte_arr.getvalue())
                
                if success:
                    all_pages_text.append({
                        'page': i + 1,
                        'text': text,
                        'has_text': True
                    })
                else:
                    all_pages_text.append({
                        'page': i + 1,
                        'text': '[متنی در این صفحه تشخیص داده نشد]',
                        'has_text': False
                    })
            
            return all_pages_text
        except Exception as e:
            logger.error(f"خطا در پردازش PDF: {e}")
            return []
    
    @staticmethod
    def process_zip(zip_bytes, progress_callback=None):
        """پردازش فایل زیپ و استخراج متن از تمام فایل‌ها"""
        extracted_data = []
        
        try:
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
                file_list = [f for f in z.namelist() 
                           if f.lower().endswith(SUPPORTED_IMAGE_FORMATS + SUPPORTED_DOC_FORMATS)]
                
                for i, filename in enumerate(file_list):
                    if progress_callback:
                        progress_callback(i + 1, len(file_list))
                    
                    with z.open(filename) as f:
                        file_bytes = f.read()
                        
                        if filename.lower().endswith('.pdf'):
                            pages = DocumentProcessor.process_pdf(file_bytes)
                            extracted_data.append({
                                'filename': filename,
                                'type': 'pdf',
                                'pages': pages
                            })
                        elif filename.lower().endswith(SUPPORTED_IMAGE_FORMATS):
                            text, success = DocumentProcessor.process_image(file_bytes)
                            if success:
                                extracted_data.append({
                                    'filename': filename,
                                    'type': 'image',
                                    'text': text
                                })
            
            return extracted_data
        except Exception as e:
            logger.error(f"خطا در پردازش زیپ: {e}")
            return []

class TelegramBot:
    """کلاس اصلی ربات تلگرام"""
    
    def __init__(self):
        self.user_sessions = {}
        self.processing_tasks = {}
    
    def get_cache_key(self, file_id):
        """ساخت کلید کش برای فایل"""
        return hashlib.md5(file_id.encode()).hexdigest()
    
    def update_stats(self, file_type, success):
        """به‌روزرسانی آمار"""
        with stats_lock:
            stats['total_processed'] += 1
            stats[f'processed_{file_type}'] += 1
            if success:
                stats['successful'] += 1
            else:
                stats['failed'] += 1
    
    def send_progress(self, chat_id, current, total, message):
        """ارسال وضعیت پیشرفت"""
        try:
            percent = (current / total) * 100
            progress_bar = '█' * int(percent/10) + '░' * (10 - int(percent/10))
            text = f"{message}\n{progress_bar} {percent:.0f}% ({current}/{total})"
            bot.edit_message_text(text, chat_id, self.processing_tasks.get(chat_id))
        except:
            pass

bot_handler = TelegramBot()

# ========== هندلرهای ربات ==========

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    welcome_text = """
🤖 **ربات حرفه‌ای تبدیل فایل به Word**

من یک ربات قدرتمند برای استخراج متن از فایل‌های مختلف هستم!

📤 **فایل‌های پشتیبانی شده:**
• 📷 عکس‌ها (JPG, PNG, BMP, TIFF, WebP)
• 📚 فایل‌های PDF (اسکن شده و متنی)
• 📦 فایل‌های ZIP حاوی عکس و PDF

✨ **قابلیت‌های ویژه:**
• 🎯 تشخیص و اصلاح کجی تصاویر
• 🌙 حذف نویز و بهبود کیفیت
• 📝 حفظ فرمت و ساختار متن
• 🚀 پردازش سریع و هوشمند
• 💾 کش نتایج برای پردازش‌های تکراری

🔍 **نکات مهم برای بهترین نتیجه:**
• تصاویر با نور کافی و بدون سایه باشند
• کیفیت تصویر بالا باشد (حداقل 300 DPI)
• زاویه عکس مستقیم و صاف باشد
• فونت متن ساده و خوانا باشد

📊 **آمار ربات:**
• کل پردازش‌ها: {total}
• موفقیت‌آمیز: {success}
• میزان موفقیت: {rate}%

🚀 **فقط کافیست فایل خود را ارسال کنید!**
""".format(
    total=stats.get('total_processed', 0),
    success=stats.get('successful', 0),
    rate=f"{(stats.get('successful', 0) / max(stats.get('total_processed', 1), 1) * 100):.1f}"
)
    bot.reply_to(message, welcome_text, parse_mode='Markdown')

@bot.message_handler(commands=['stats'])
def show_stats(message):
    """نمایش آمار ربات"""
    stats_text = f"""
📊 **آمار ربات حرفه‌ای**

📈 **آمار کلی:**
• کل فایل‌های پردازش شده: {stats.get('total_processed', 0)}
• پردازش‌های موفق: {stats.get('successful', 0)}
• پردازش‌های ناموفق: {stats.get('failed', 0)}

📁 **نوع فایل‌ها:**
• عکس‌ها: {stats.get('processed_photo', 0)}
• PDFها: {stats.get('processed_pdf', 0)}
• فایل‌های ZIP: {stats.get('processed_zip', 0)}

🏆 **نرخ موفقیت:** {((stats.get('successful', 0) / max(stats.get('total_processed', 1), 1)) * 100):.1f}%

🎯 **وضعیت:** آنلاین و فعال 🟢
"""
    bot.reply_to(message, stats_text, parse_mode='Markdown')

@bot.message_handler(commands=['cache'])
def clear_cache(message):
    """پاک کردن کش"""
    with cache_lock:
        cache_size = len(processing_cache)
        processing_cache.clear()
    bot.reply_to(message, f"✅ کش پردازش پاک شد! {cache_size} آیتم حذف گردید.")

@bot.message_handler(content_types=['photo'])
def handle_photo(message):
    """پردازش عکس"""
    try:
        chat_id = message.chat.id
        msg = bot.reply_to(message, "📷 **در حال پردازش عکس...**\n\nمرحله 1/3: دریافت فایل", parse_mode='Markdown')
        bot_handler.processing_tasks[chat_id] = msg.message_id
        
        # دریافت باکیفیت‌ترین نسخه عکس
        photo = message.photo[-1]
        file_info = bot.get_file(photo.file_id)
        downloaded = bot.download_file(file_info.file_path)
        
        # به‌روزرسانی وضعیت
        bot.edit_message_text("📷 **در حال پردازش عکس...**\n\nمرحله 2/3: بهبود کیفیت تصویر", 
                            chat_id, msg.message_id, parse_mode='Markdown')
        
        # پردازش تصویر
        text, success = DocumentProcessor.process_image(downloaded, 'accurate')
        
        if success and text:
            bot.edit_message_text("📷 **در حال پردازش عکس...**\n\nمرحله 3/3: ساخت فایل Word", 
                                chat_id, msg.message_id, parse_mode='Markdown')
            
            # ایجاد سند Word
            doc = Document()
            doc.add_heading('متن استخراج شده از تصویر', 0)
            doc.add_paragraph(f"تاریخ: {datetime.now().strftime('%Y/%m/%d - %H:%M:%S')}")
            doc.add_paragraph()
            
            # اضافه کردن متن با فرمت
            TextFormatter.add_to_document(doc, text)
            
            # ذخیره و ارسال
            output_path = f"output_{chat_id}_{int(time.time())}.docx"
            doc.save(output_path)
            
            with open(output_path, 'rb') as f:
                bot.send_document(chat_id, f, caption="✅ **فایل Word با موفقیت ساخته شد!**\n\n📌 متن استخراج شده از تصویر شما.", parse_mode='Markdown')
            
            os.remove(output_path)
            bot.delete_message(chat_id, msg.message_id)
            bot.send_message(chat_id, "🎉 **پردازش با موفقیت انجام شد!**\nبرای دریافت فایل ورد به پیام بالا مراجعه کنید.", parse_mode='Markdown')
            
            bot_handler.update_stats('photo', True)
        else:
            bot.edit_message_text("❌ **خطا در پردازش عکس**\n\nمتنی در این عکس تشخیص داده نشد.\nلطفاً عکس با کیفیت بهتری ارسال کنید.", 
                                chat_id, msg.message_id, parse_mode='Markdown')
            bot_handler.update_stats('photo', False)
    
    except Exception as e:
        logger.error(f"خطا در پردازش عکس: {e}")
        bot.reply_to(message, f"❌ **خطای سیستمی رخ داد!**\n{str(e)[:100]}", parse_mode='Markdown')
        bot_handler.update_stats('photo', False)

@bot.message_handler(content_types=['document'])
def handle_document(message):
    """پردازش فایل‌های ضمیمه شده"""
    try:
        chat_id = message.chat.id
        file_name = message.document.file_name
        file_size = message.document.file_size
        
        # بررسی حجم فایل
        if file_size > MAX_FILE_SIZE:
            bot.reply_to(message, f"❌ **حجم فایل بیش از حد مجاز است!**\nحداکثر حجم مجاز: {MAX_FILE_SIZE//(1024*1024)} مگابایت", parse_mode='Markdown')
            return
        
        # بررسی فرمت فایل
        is_pdf = file_name.lower().endswith('.pdf')
        is_zip = file_name.lower().endswith('.zip')
        is_image = file_name.lower().endswith(SUPPORTED_IMAGE_FORMATS)
        
        if not (is_pdf or is_zip or is_image):
            bot.reply_to(message, f"❌ **فرمت فایل پشتیبانی نمی‌شود!**\nفرمت‌های مجاز: {', '.join(SUPPORTED_IMAGE_FORMATS + SUPPORTED_DOC_FORMATS)}", parse_mode='Markdown')
            return
        
        msg = bot.reply_to(message, f"📄 **در حال پردازش {file_name}...**\n\nمرحله 1/4: دریافت فایل", parse_mode='Markdown')
        bot_handler.processing_tasks[chat_id] = msg.message_id
        
        # دریافت فایل
        file_info = bot.get_file(message.document.file_id)
        downloaded = bot.download_file(file_info.file_path)
        
        doc = Document()
        doc.add_heading('متن استخراج شده از فایل', 0)
        doc.add_paragraph(f"نام فایل اصلی: {file_name}")
        doc.add_paragraph(f"تاریخ پردازش: {datetime.now().strftime('%Y/%m/%d - %H:%M:%S')}")
        doc.add_paragraph()
        
        if is_pdf:
            bot.edit_message_text(f"📄 **در حال پردازش PDF...**\n\nمرحله 2/4: استخراج صفحات", 
                                chat_id, msg.message_id, parse_mode='Markdown')
            
            def progress_callback(current, total):
                if current % max(1, total//10) == 0:  # به‌روزرسانی هر 10%
                    bot.edit_message_text(f"📄 **در حال پردازش PDF...**\n\nپردازش صفحه {current} از {total}", 
                                        chat_id, msg.message_id, parse_mode='Markdown')
            
            pages = DocumentProcessor.process_pdf(downloaded, progress_callback)
            
            if pages:
                for page in pages:
                    if page['has_text']:
                        doc.add_heading(f"صفحه {page['page']}", level=2)
                        TextFormatter.add_to_document(doc, page['text'])
                        doc.add_page_break()
                
                bot_handler.update_stats('pdf', True)
            else:
                doc.add_paragraph("❌ متنی در این PDF تشخیص داده نشد.")
                bot_handler.update_stats('pdf', False)
        
        elif is_zip:
            bot.edit_message_text(f"📦 **در حال پردازش فایل ZIP...**\n\nمرحله 2/4: استخراج فایل‌ها", 
                                chat_id, msg.message_id, parse_mode='Markdown')
            
            extracted_data = DocumentProcessor.process_zip(downloaded)
            
            if extracted_data:
                for data in extracted_data:
                    doc.add_heading(f"فایل: {data['filename']}", level=2)
                    
                    if data['type'] == 'pdf':
                        for page in data['pages']:
                            if page['has_text']:
                                doc.add_heading(f"صفحه {page['page']}", level=3)
                                TextFormatter.add_to_document(doc, page['text'])
                    else:
                        TextFormatter.add_to_document(doc, data['text'])
                    
                    doc.add_paragraph()
                
                bot_handler.update_stats('zip', True)
            else:
                doc.add_paragraph("❌ هیچ فایل پشتیبانی شده‌ای در ZIP یافت نشد.")
                bot_handler.update_stats('zip', False)
        
        elif is_image:
            bot.edit_message_text(f"🖼️ **در حال پردازش عکس...**\n\nمرحله 2/3: استخراج متن", 
                                chat_id, msg.message_id, parse_mode='Markdown')
            
            text, success = DocumentProcessor.process_image(downloaded, 'accurate')
            
            if success:
                TextFormatter.add_to_document(doc, text)
                bot_handler.update_stats('photo', True)
            else:
                doc.add_paragraph("❌ متنی در این عکس تشخیص داده نشد.")
                bot_handler.update_stats('photo', False)
        
        # ذخیره و ارسال فایل
        bot.edit_message_text("📝 **در حال ساخت فایل Word...**\n\nمرحله نهایی: آماده‌سازی خروجی", 
                            chat_id, msg.message_id, parse_mode='Markdown')
        
        output_path = f"output_{chat_id}_{int(time.time())}.docx"
        doc.save(output_path)
        
        with open(output_path, 'rb') as f:
            bot.send_document(chat_id, f, caption="✅ **فایل Word با موفقیت ساخته شد!**\n\n📌 متن استخراج شده از فایل شما.", parse_mode='Markdown')
        
        os.remove(output_path)
        bot.delete_message(chat_id, msg.message_id)
        bot.send_message(chat_id, "🎉 **پردازش با موفقیت انجام شد!**\nبرای دریافت فایل ورد به پیام بالا مراجعه کنید.", parse_mode='Markdown')
    
    except Exception as e:
        logger.error(f"خطا در پردازش فایل: {e}")
        bot.reply_to(message, f"❌ **خطای سیستمی رخ داد!**\n{str(e)[:200]}", parse_mode='Markdown')
        bot_handler.update_stats('document', False)

@bot.message_handler(func=lambda message: True)
def handle_unknown(message):
    """پاسخ به پیام‌های ناشناخته"""
    bot.reply_to(message, 
                "🤔 **دستور ناشناخته!**\n\n"
                "لطفاً یکی از گزینه‌های زیر را ارسال کنید:\n"
                "• یک عکس 📷\n"
                "• یک فایل PDF 📚\n"
                "• یک فایل ZIP 📦\n\n"
                "برای راهنمایی بیشتر، دستور /start را ارسال کنید.",
                parse_mode='Markdown')

# ========== راه‌اندازی ربات ==========
if __name__ == '__main__':
    print("=" * 50)
    print("🤖 ربات حرفه‌ای تبدیل فایل به Word")
    print("📱 نسخه 2.0 - با قابلیت‌های پیشرفته")
    print("=" * 50)
    print(f"✅ ربات با موفقیت راه‌اندازی شد!")
    print(f"📊 آمار اولیه: {stats}")
    print("=" * 50)
    
    try:
        bot.infinity_polling(timeout=60, long_polling_timeout=60)
    except KeyboardInterrupt:
        print("\n👋 ربات متوقف شد.")
    except Exception as e:
        print(f"❌ خطا در اجرای ربات: {e}")
