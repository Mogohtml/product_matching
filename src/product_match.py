import pandas as pd
import re
import random
from datetime import datetime

# Словарь для перевода цветов на английский язык
color_translation = {
    "серебристый": "Silver",
    "серый космос": "Space Gray",
    "черный": "Black",
    "белый": "White",
    "розовый": "Pink",
    "синий": "Blue",
    "зеленый": "Green",
    "красный": "Red",
    "фиолетовый": "Purple",
    "оранжевый": "Orange",
    "желтый": "Yellow",
    "золотой": "Gold",
}

# Функция для создания словаря характеристик товара
def create_item_dict(item_name, characteristics):
    words = set(item_name.lower().split() + characteristics.lower().split())
    return words

# Функция для извлечения информации о продукте из строки
def extract_product_info(row):
    if isinstance(row, str):
        # Регулярное выражение для извлечения артикула, названия, цены и даты
        match = re.match(r'(\d+)\s+(.+?)\s+(\d{5})\s*(.*)', row)
        if match:
            article, name, price, extra = match.groups()
            # Определяем статус на основе наличия текста после цены
            status = "В наличии" if extra else "Нет в наличии"
            # Извлекаем цвет из названия
            color_match = re.search(r'\b(\w+)\s+JP\s+\d{5}', name)
            color = color_match.group(1).lower() if color_match else None
            # Извлекаем все возможные даты
            dates = re.findall(r'\[(\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2})\]', row)
            date_str = max(dates, key=lambda d: datetime.strptime(d, '%d.%m.%Y %H:%M')) if dates else None
            return article, name.strip(), price, status, color, date_str
    return None, None, None, None, None, None

# Функция для генерации случайного артикула
def generate_random_article(last_article):
    min_article = max(14000, last_article - 20)
    max_article = min(17000, last_article + 20)
    article = random.randint(min_article, max_article)
    return article

# Функция для сопоставления товаров из магазина с товарами поставщиков
def match_items(shop_df, supplier_df):
    matched_items = []
    last_article = 14000  # Инициализируем начальное значение

    for index, shop_row in shop_df.iterrows():
        shop_name = shop_row['Наименование']
        characteristics = f"{shop_row['Производитель']} {shop_row['Модель']} {shop_row['Цвет']} {shop_row['Встроенная память']}"
        shop_dict = create_item_dict(shop_name, characteristics)

        matched_suppliers = []
        for _, supplier_row in supplier_df.iterrows():
            supplier_name = supplier_row['Название']
            if pd.isna(supplier_name):
                continue
            supplier_words = set(supplier_name.lower().split())

            if supplier_words & shop_dict:
                price = supplier_row['Цена']
                supplier = supplier_row['поставщик']
                status = supplier_row.get('Статус', '')
                color = supplier_row.get('Цвет', '')
                date_str = supplier_row.get('Дата', '')
                matched_suppliers.append((price, supplier, status, color, date_str))

        days_in_stock = None  # Инициализируем переменную
        if matched_suppliers:
            min_price = min(int(price) for price, _, _, _, _ in matched_suppliers)
            recommended_prices = {
                "Рекомендуемая Саратов": round(min_price * 1.10),
                "Рекомендуемая Воронеж": round(min_price * 1.10),
                "Рекомендуемая Липецк": round(min_price * 1.10),
            }
            # Определяем статус
            status = matched_suppliers[0][2] if matched_suppliers[0][2] else "Нет в наличии"
            color = matched_suppliers[0][3]
            english_color = color_translation.get(color, "")
            # Определяем дату
            date_str = matched_suppliers[0][4]
            if date_str and status == "В наличии":
                date_str = date_str.strip('[]')  # Убираем квадратные скобки
                days_in_stock = (datetime.now() - datetime.strptime(date_str, '%d.%m.%Y %H:%M')).days
        else:
            recommended_prices = {
                "Рекомендуемая Саратов": None,
                "Рекомендуемая Воронеж": None,
                "Рекомендуемая Липецк": None,
            }
            status = "Нет в наличии"
            english_color = ""

        # Формируем новое название с иерархией по памяти
        memory = shop_row.get('Встроенная память', '')
        new_name = f"{shop_name} ({english_color}) {memory} nano SIM + eSIM"

        # Извлекаем или генерируем артикул
        article = shop_row.get('Артикул', '')
        if not article:
            article = generate_random_article(last_article)
            last_article = article

        matched_items.append({
            "Артикул": article,
            "Внешний код": shop_row['Внешний код'],
            "Новое название": new_name,
            "Статус": status,
            "Есть на складе?": "Есть" if status == "В наличии" else "Нет",
            "Дней на складе": days_in_stock,
            "Себестоимость": shop_row.get('Себестоимость', ''),
            "Кол-во": shop_row.get('Кол-во', ''),
            "Продаж за 30 дней": shop_row.get('Продаж за 30 дней', ''),
            "Продаж за неделю": shop_row.get('Продаж за неделю', ''),
            "Дата последней продажи": shop_row.get('Дата последней продажи', ''),
            "Цена последней продажи": round(float(shop_row.get('Цена последней продажи', ''))) if shop_row.get('Цена последней продажи', '') else None,
            "Заказано": shop_row.get('Заказано', ''),
            "Цена: Саратов": round(float(shop_row.get('Цена: Саратов', ''))) if shop_row.get('Цена: Саратов', '') else None,
            "Цена: Воронеж": round(float(shop_row.get('Цена: Воронеж', ''))) if shop_row.get('Цена: Воронеж', '') else None,
            "Цена: Липецк": round(float(shop_row.get('Цена: Липецк', ''))) if shop_row.get('Цена: Липецк', '') else None,
            **recommended_prices,
            "Поставщик 1": matched_suppliers[0][1] if len(matched_suppliers) > 0 else '',
            "Цена поставщика 1": round(float(matched_suppliers[0][0])) if len(matched_suppliers) > 0 else '',
            "Поставщик 2": matched_suppliers[1][1] if len(matched_suppliers) > 1 else '',
            "Цена поставщика 2": round(float(matched_suppliers[1][0])) if len(matched_suppliers) > 1 else '',
            "Поставщик 3": matched_suppliers[2][1] if len(matched_suppliers) > 2 else '',
            "Цена поставщика 3": round(float(matched_suppliers[2][0])) if len(matched_suppliers) > 2 else '',
        })

    return matched_items

def main(shop_file, supplier_file):
    shop_df = pd.read_excel(shop_file)
    supplier_df = pd.read_excel(supplier_file)

    # Извлекаем информацию о продуктах из столбца 'прайс'
    supplier_df[['Артикул', 'Название', 'Цена', 'Статус', 'Цвет', 'Дата']] = supplier_df['прайс'].apply(lambda x: pd.Series(extract_product_info(x)))
    supplier_df = supplier_df.dropna(subset=['Название', 'Цена']).reset_index(drop=True)

    # Преобразуем цены в целочисленный формат
    supplier_df['Цена'] = pd.to_numeric(supplier_df['Цена'], downcast='integer', errors='coerce')

    matched_items = match_items(shop_df, supplier_df)

    result_df = pd.DataFrame(matched_items)

    result_df = result_df[[
        "Артикул", "Внешний код", "Новое название", "Статус", "Есть на складе?", "Дней на складе",
        "Себестоимость", "Кол-во", "Продаж за 30 дней", "Продаж за неделю", "Дата последней продажи",
        "Цена последней продажи", "Заказано", "Цена: Саратов", "Цена: Воронеж", "Цена: Липецк",
        "Рекомендуемая Саратов", "Рекомендуемая Воронеж", "Рекомендуемая Липецк",
        "Поставщик 1", "Цена поставщика 1", "Поставщик 2", "Цена поставщика 2", "Поставщик 3", "Цена поставщика 3"
    ]]

    result_df.to_csv("matched_products.csv", index=False)
    print("Сопоставление завершено. Результаты сохранены в файл 'matched_products.csv'.")

if __name__ == "__main__":
    shop_file = r"D:\Учеба\shop_items.xlsx"
    supplier_file = r"D:\Учеба\supplier_items.xlsx"
    main(shop_file, supplier_file)
