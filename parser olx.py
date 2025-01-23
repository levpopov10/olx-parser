import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
import time

# Исходные данные с брендами
brands_data = """
Acura98
Audi1 608
BMW1 696
Chevrolet493
Chrysler95
Dodge119
Fiat261
Ford1 390
Honda426
Hyundai876
Infiniti165
Jeep511
Kia779
Land Rover296
Lexus346
Lincoln85
Mazda706
Mercedes-Benz1 490
Mitsubishi925
Nissan1 328
Porsche149
Subaru370
Tesla435
Toyota968
Volkswagen2 237
Volvo320
drugoy-brend
"""

# Обработка данных
brand_list = brands_data.strip().split("\n")
cleaned_brands = []
for brand in brand_list:
    cleaned_brand = ''.join(filter(str.isalpha, brand)).replace(" ", "-").lower()
    link = f"https://www.olx.ua/uk/zapchasti-dlya-transporta/transport-na-zapchasti-avtorazborki/{cleaned_brand}/?currency=UAH"
    cleaned_brands.append(link)

# Список для сохранения данных
data = []
# max_ads_to_process = 5  # Лимит на количество обрабатываемых объявлений

# Вывод ссылок
for index, link in enumerate(cleaned_brands):
    # URL главной страницы
    url = link
    PHONE_URL = "https://www.olx.ua/api/v1/phones/"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36'}

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, "html.parser")
        links = soup.find_all("a", class_="css-qo0cxu")

        processed_links = set()  # Для хранения уникальных ссылок объявлений
        
        for ad_index, link1 in enumerate(links):
            # if ad_index >= max_ads_to_process:  # Проверка лимита на объявления
            #     print("Достигнуто максимальное количество обрабатываемых объявлений в этом бренде.")
            #     break
            
            href = link1.get('href')
            if href.startswith('/'):
                href = "https://www.olx.ua" + href
            
            if href in processed_links:  # Проверка на дубликаты
                print(f"Ссылка {href} уже обработана, пропускаем.")
                continue
            
            processed_links.add(href)  # Добавляем ссылку в обработанные
            
            print(f"Переход к ссылке: {href}")
            ad_response = requests.get(href, headers=headers)

            if ad_response.status_code == 200:
                ad_soup = BeautifulSoup(ad_response.content, "html.parser")
                
                # Извлечение текста из title
                title_tag = ad_soup.title
                if title_tag:
                    title_text = title_tag.text.strip()
                    name = title_text.split(":")[0].strip() if ":" in title_text else title_text
                    price_part = title_text.split(":")[1].split("-")[0].strip() if ":" in title_text and "-" in title_text else "Цена не указана"
                    
                    # Извлечение описания
                    description_div = ad_soup.find("div", class_="css-1o924a9")
                    description = description_div.text.strip() if description_div else "Описание не найдено"
                    
                    # Извлечение номера телефона
                    ad_id = href.split("/")[-1].split("-")[-1]
                    phone_response = requests.get(f"{PHONE_URL}{ad_id}", headers=headers)

                    if phone_response.status_code == 200:
                        phone_data = phone_response.json()
                        phone_number = phone_data.get("phone_number", "Не указано")
                    else:
                        phone_number = "Не удалось получить номер"
                    
                    # Извлечение текста из h4
                    h4_tags = ad_soup.find_all("h4", class_="css-lyp0yk")
                    h4_texts = [h4_tag.text.strip() for h4_tag in h4_tags] if h4_tags else ["Не найдено"]
                    
                    # Сохранение данных
                    brand_name = brand_list[index].rstrip('0123456789 ')
                    data.append({
                        "Название": name,
                        "Цена": price_part,
                        "Телефон": phone_number,
                        "Описание": description,
                        "Ссылка": href,
                        "Бренд": brand_name,
                        "h4_тексты": h4_texts[0],
                    })
                else:
                    print("Title не найден.")
            else:
                print(f"Не удалось загрузить страницу. Код ошибки: {ad_response.status_code}")
            time.sleep(1)
    else:
        print(f"Не удалось загрузить главную страницу. Код ошибки: {response.status_code}")

# Сохранение данных в Excel после завершения обработки
if data:
    df = pd.DataFrame(data)
    excel_file = "olx_ads.xlsx"
    df.to_excel(excel_file, index=False)

    # Изменение ширины колонок
    wb = load_workbook(excel_file)
    ws = wb.active
    column_widths = {
        "A": 25,  # Название
        "B": 15,  # Цена
        "C": 15,  # Телефон
        "D": 30,  # Описание
        "E": 30,  # Ссылка
        "F": 20,  # Бренд
        "G": 30,  # h4_тексты
    }

    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    wb.save(excel_file)  # Сохраняем изменения
    wb.close()

    print(f"Данные успешно сохранены в файл '{excel_file}' с измененными размерами ячеек.")
else:
    print("Нет данных для записи в Excel.")
