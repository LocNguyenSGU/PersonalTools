import time

import openpyxl
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl.drawing.image import Image
import requests
from io import BytesIO

BASE_URL = "https://itviec.com/companies?page={}"

def scrape_itviec():
    options = Options()
    # options.add_argument("--headless")  # Bỏ headless để kiểm tra
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    company_data = []

    for page in range(1, 31):  # Test 2 trang trướcs
        url = BASE_URL.format(page)
        driver.get(url)
        time.sleep(7)  # Chờ trang tải xong

        companies = driver.find_elements(By.CLASS_NAME, "featured-company")
        print(f"Page {page}: Found {len(companies)} companies")  # Kiểm tra số lượng

        for company in companies:
            try:
                name = company.find_element(By.CLASS_NAME, "company__name").text if company.find_elements(By.CLASS_NAME, "company__name") else "N/A"
                location = company.find_element(By.CLASS_NAME, "company__footer-city").text if company.find_elements(By.CLASS_NAME, "company__footer-city") else "N/A"
                jobs_elements = company.find_elements(By.CLASS_NAME, "d-flex")
                jobs = "N/A"
                for el in jobs_elements:
                    if "job" in el.text:
                        jobs = el.text
                        break

                reviews = company.find_element(By.CLASS_NAME, "company__footer-reviews").text if company.find_elements(By.CLASS_NAME, "company__footer-reviews") else "N/A"
                rating = company.find_element(By.CLASS_NAME, "company__star-rate").text if company.find_elements(By.CLASS_NAME, "company__star-rate") else "N/A"
                logo_img = company.find_element(By.CLASS_NAME, "company__logo").find_element(By.TAG_NAME, "img").get_attribute("src") if company.find_elements(By.CLASS_NAME, "company__logo") else "N/A"
                company_url = company.find_element(By.TAG_NAME, "a").get_attribute("href") if company.find_elements(By.TAG_NAME, "a") else "N/A"

                company_info = {
                    "Name": name,
                    "Location": location,
                    "Jobs": jobs,
                    "Reviews": reviews,
                    "Rating": rating,
                    "Logo URL": logo_img,
                    "Company URL": company_url
                }
                print(company_info)  # Kiểm tra dữ liệu từng công ty
                company_data.append(company_info)

            except Exception as e:
                print("Error:", e)

    driver.quit()

    if company_data:
        df = pd.DataFrame(company_data)
        df.to_excel("itviec_companies_selenium.xlsx", index=False)
        print("Data saved to itviec_companies_selenium.xlsx")
    else:
        print("No data collected!")


def insert_images_into_excel(file_path, image_column="F", output_column="H",
                             output_file="itviec_companies_with_images.xlsx"):
    """
    Chèn hình ảnh từ URL vào file Excel.

    Parameters:
    - file_path (str): Đường dẫn file Excel.
    - image_column (str): Cột chứa URL ảnh.
    - output_column (str): Cột sẽ hiển thị ảnh.
    - output_file (str): Tên file Excel đầu ra.
    """

    # Mở file Excel
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active  # Chọn sheet đầu tiên

    # Lặp qua từng hàng để chèn ảnh
    for row in range(2, ws.max_row + 1):  # Bỏ qua hàng tiêu đề
        img_url = ws[f"{image_column}{row}"].value  # Lấy URL ảnh
        if img_url:  # Nếu có URL
            try:
                # Tải ảnh từ URL
                response = requests.get(img_url)
                img = Image(BytesIO(response.content))  # Đọc dữ liệu ảnh

                # Resize ảnh (có thể chỉnh sửa kích thước tùy ý)
                img.width = 80  # Độ rộng (px)
                img.height = 80  # Chiều cao (px)

                # Chèn ảnh vào cột chỉ định
                ws.add_image(img, f"{output_column}{row}")

            except Exception as e:
                print(f"❌ Lỗi khi tải ảnh ở hàng {row}: {e}")

    # Lưu file Excel mới
    wb.save(output_file)
    print(f"✅ Đã chèn ảnh và lưu vào {output_file}")


# import requests
#
# def create_deck(deck_name):
#     response = requests.post("http://localhost:8765", json={
#         "action": "createDeck",
#         "version": 6,
#         "params": {
#             "deck": deck_name
#         }
#     })
#     return response.json()
#
# def add_note(front, back, deck="Default", model="Basic", tag="auto-added"):
#     response = requests.post("http://localhost:8765", json={
#         "action": "addNote",
#         "version": 6,
#         "params": {
#             "note": {
#                 "deckName": deck,
#                 "modelName": model,
#                 "fields": {
#                     "Front": front,
#                     "Back": back
#                 },
#                 "options": {
#                     "allowDuplicate": False
#                 },
#                 "tags": [tag]
#             }
#         }
#     })
#     return response.json()
#
# def import_txt_to_anki(txt_file):
#     created_decks = set()
#
#     with open(txt_file, 'r', encoding='utf-8') as f:
#         for line in f:
#             if ':::' in line:
#                 parts = line.strip().split(':::')
#                 if len(parts) >= 2:
#                     question = parts[0].strip()
#                     answer = parts[1].strip()
#                     example = parts[2].strip() if len(parts) >= 3 else ""
#                     topic = parts[3].strip() if len(parts) >= 4 else "General"
#
#                     # Gộp câu trả lời với ví dụ xuống dòng
#                     full_back = f"{answer}<br><br><i>Example:</i><br>{example}" if example else answer
#
#                     # Tạo deck nếu chưa có
#                     if topic not in created_decks:
#                         create_deck(topic)
#                         created_decks.add(topic)
#
#                     # Thêm thẻ vào deck đúng với chủ đề
#                     result = add_note(question, full_back, deck=topic, model="Basic", tag=topic)
#                     print(f"✔️ Thêm: {question} vào deck [{topic}] -> {result}")
#
# if __name__ == "__main__":
#     import_txt_to_anki("cards.txt")


    # scrape_itviec()
    # insert_images_into_excel(file_path="itviec_companies_selenium.xlsx");

import requests
from gtts import gTTS
import os
import hashlib

def generate_audio(word, folder="media"):
    if not os.path.exists(folder):
        os.makedirs(folder)

    # Thay thế khoảng trắng bằng dấu gạch dưới
    word_modified = word.replace(" ", "_")

    # Tạo tên file gồm hash của từ và từ đã thay thế khoảng trắng
    file_name = f"{word_modified}_{hashlib.md5(word_modified.encode()).hexdigest()}.mp3"
    audio_path = os.path.join(folder, file_name)

    if not os.path.exists(audio_path):
        tts = gTTS(text=word, lang='en')
        tts.save(audio_path)

    return audio_path

import base64

def add_note(front, back, example, deck, topic_tag, audio_path=None, image_path=None):
    media = []
    if audio_path:
        with open(audio_path, "rb") as f:
            audio_data = base64.b64encode(f.read()).decode('utf-8')  # Mã hóa Base64
            media.append({
                "filename": os.path.basename(audio_path),
                "data": audio_data,
                "fields": ["Back"]
            })
    if image_path:
        with open(image_path, "rb") as f:
            image_data = base64.b64encode(f.read()).decode('utf-8')  # Mã hóa Base64
            media.append({
                "filename": os.path.basename(image_path),
                "data": image_data,
                "fields": ["Back"]
            })

    back_content = f"{back}<br><br><i>Example:</i><br>{example}"
    if image_path:
        back_content += f'<br><img src="{os.path.basename(image_path)}">'
    if audio_path:
        back_content += f'<br>[sound:{os.path.basename(audio_path)}]'

    payload = {
        "action": "addNote",
        "version": 6,
        "params": {
            "note": {
                "deckName": deck,
                "modelName": "Basic",
                "fields": {
                    "Front": front,
                    "Back": back_content
                },
                "options": {
                    "allowDuplicate": False
                },
                "tags": ["auto-added", topic_tag]
            }
        }
    }

    if media:
        for m in media:
            requests.post("http://localhost:8765", json={
                "action": "storeMediaFile",
                "version": 6,
                "params": {
                    "filename": m["filename"],
                    "data": m["data"]
                }
            })

    return requests.post("http://localhost:8765", json=payload).json()

# Cố định tên deck
import logging

# Cấu hình logging để in ra console với thời gian
logging.basicConfig(
    filename="import_log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def create_deck(deck_name):
    response = requests.post("http://localhost:8765", json={
        "action": "createDeck",
        "version": 6,
        "params": {
            "deck": deck_name
        }
    })
    return response.json()


def import_txt_to_anki(txt_file):
    created_decks = set()  # Sử dụng set để lưu các deck đã tạo

    with open(txt_file, 'r', encoding='utf-8') as f:
        for line in f:
            parts = line.strip().split(':::')
            if len(parts) == 4:
                vi, en, example, topic = [p.strip() for p in parts]

                # Tạo deck mới cho từng topic nếu chưa có
                if topic not in created_decks:
                    create_deck(topic)  # Gọi hàm tạo deck
                    created_decks.add(topic)

                # Thêm thẻ vào deck tương ứng
                audio_path = generate_audio(en)
                result = add_note(vi, en, example, topic, topic, audio_path, None)  # Tạo thẻ cho từng deck riêng biệt
                print(f"✅ Đã thêm [{topic}] vào deck [{topic}]: {vi} - {en}")
                logging.info(f"✅ Đã thêm [{topic}] vào deck [{topic}]: {vi} - {en}")
            else:
                print(f"❌ Lỗi format dòng: {line.strip()}")
                logging.error(f"❌ Lỗi format dòng: {line.strip()}")

    # ==== 📌 Chạy script ====
if __name__ == "__main__":
    import_txt_to_anki("cards.txt")