import telebot
from docx import Document
from PIL import Image
import pytesseract
import io
import os
from pdf2image import convert_from_bytes
import zipfile

TOKEN = os.environ.get("TOKEN")
bot = telebot.TeleBot(TOKEN)

@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(message, "سلام! هر فایل عکس، PDF یا زیپ رو بفرست تا به Word قابل سرچ تبدیلش کنم.")

@bot.message_handler(content_types=['photo', 'document'])
def handle_files(message):
    chat_id = message.chat.id
    bot.send_message(chat_id, "⚙️ در حال پردازش...")
    
    doc = Document()
    text_content = []
    
    if message.photo:
        file_info = bot.get_file(message.photo[-1].file_id)
        downloaded = bot.download_file(file_info.file_path)
        img = Image.open(io.BytesIO(downloaded))
        text = pytesseract.image_to_string(img, lang='eng+fas')
        text_content.append(text)
    
    elif message.document:
        file_info = bot.get_file(message.document.file_id)
        downloaded = bot.download_file(file_info.file_path)
        file_name = message.document.file_name
        
        if file_name.endswith('.pdf'):
            images = convert_from_bytes(downloaded)
            for img in images:
                text = pytesseract.image_to_string(img, lang='eng+fas')
                text_content.append(text)
                
        elif file_name.endswith(('.png', '.jpg', '.jpeg')):
            img = Image.open(io.BytesIO(downloaded))
            text = pytesseract.image_to_string(img, lang='eng+fas')
            text_content.append(text)
            
        elif file_name.endswith('.zip'):
            with zipfile.ZipFile(io.BytesIO(downloaded)) as z:
                for file_in_zip in z.namelist():
                    if file_in_zip.endswith(('.png', '.jpg', '.jpeg')):
                        with z.open(file_in_zip) as img_file:
                            img = Image.open(img_file)
                            text = pytesseract.image_to_string(img, lang='eng+fas')
                            text_content.append(f"\n--- فایل: {file_in_zip} ---\n{text}")
    
    for paragraph in text_content:
        doc.add_paragraph(paragraph)
    
    output_path = f"output_{chat_id}.docx"
    doc.save(output_path)
    
    with open(output_path, 'rb') as f:
        bot.send_document(chat_id, f, caption="✅ فایل Word قابل سرچ آماده است!")
    
    os.remove(output_path)
    bot.send_message(chat_id, "🎉 انجام شد!")

print("ربات روشن شد...")
bot.infinity_polling()
