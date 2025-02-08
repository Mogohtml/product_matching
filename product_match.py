import pandas as pd
import re

def create_item_dict(item_name, characteristics):
    words = set(item_name.lower().split() + characteristics.lower().split())
    return words

def extract_product_info(row):
    if isinstance(row, str):
        match = re.match(r'(.+?)\s+(\d+)', row)
        if match:
            return match.groups()
    return None, None

def match_items(shop_df, supplier_df):
    matched_items = []

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
                matched_suppliers.append((price, supplier))

        # Calculate recommended prices
        if matched_suppliers:
            min_price = min(float(price) for price, _ in matched_suppliers)  # Convert price to float
            recommended_prices = {
                "Рекомендуемая Саратов": min_price * 1.10,
                "Рекомендуемая Воронеж": min_price * 1.10,
                "Рекомендуемая Липецк": min_price * 1.10,
            }
        else:
            recommended_prices = {
                "Рекомендуемая Саратов": None,
                "Рекомендуемая Воронеж": None,
                "Рекомендуемая Липецк": None,
            }

        matched_items.append({
            "Артикул": shop_row.get('Артикул', ''),
            "Внешний код": shop_row['Внешний код'],
            "Новое название": shop_name,
            "Статус": shop_row.get('Статус', ''),
            "Есть на складе?": shop_row.get('Есть на складе?', ''),
            "Дней на складе": shop_row.get('Дней на складе', ''),
            "Себестоимость": shop_row.get('Себестоимость', ''),
            "Кол-во": shop_row.get('Кол-во', ''),
            "Продаж за 30 дней": shop_row.get('Продаж за 30 дней', ''),
            "Продаж за неделю": shop_row.get('Продаж за неделю', ''),
            "Дата последней продажи": shop_row.get('Дата последней продажи', ''),
            "Цена последней продажи": shop_row.get('Цена последней продажи', ''),
            "Заказано": shop_row.get('Заказано', ''),
            "Цена: Саратов": shop_row.get('Цена: Саратов', ''),
            "Цена: Воронеж": shop_row.get('Цена: Воронеж', ''),
            "Цена: Липецк": shop_row.get('Цена: Липецк', ''),
            **recommended_prices,
            "Поставщик 1": matched_suppliers[0][1] if len(matched_suppliers) > 0 else '',
            "Цена поставщика 1": matched_suppliers[0][0] if len(matched_suppliers) > 0 else '',
            "Поставщик 2": matched_suppliers[1][1] if len(matched_suppliers) > 1 else '',
            "Цена поставщика 2": matched_suppliers[1][0] if len(matched_suppliers) > 1 else '',
            "Поставщик 3": matched_suppliers[2][1] if len(matched_suppliers) > 2 else '',
            "Цена поставщика 3": matched_suppliers[2][0] if len(matched_suppliers) > 2 else '',
        })

    return matched_items

def main(shop_file, supplier_file):
    shop_df = pd.read_excel(shop_file)
    supplier_df = pd.read_excel(supplier_file)

    supplier_df[['Название', 'Цена']] = supplier_df['прайс'].apply(lambda x: pd.Series(extract_product_info(x)))
    supplier_df = supplier_df.dropna(subset=['Название', 'Цена']).reset_index(drop=True)

    # Convert 'Цена' column to numeric
    supplier_df['Цена'] = pd.to_numeric(supplier_df['Цена'], errors='coerce')

    matched_items = match_items(shop_df, supplier_df)

    # Преобразование списка сопоставленных товаров в DataFrame
    result_df = pd.DataFrame(matched_items)

    # Reorder columns to match the expected format
    result_df = result_df[[
        "Артикул", "Внешний код", "Новое название", "Статус", "Есть на складе?", "Дней на складе",
        "Себестоимость", "Кол-во", "Продаж за 30 дней", "Продаж за неделю", "Дата последней продажи",
        "Цена последней продажи", "Заказано", "Цена: Саратов", "Цена: Воронеж", "Цена: Липецк",
        "Рекомендуемая Саратов", "Рекомендуемая Воронеж", "Рекомендуемая Липецк",
        "Поставщик 1", "Цена поставщика 1", "Поставщик 2", "Цена поставщика 2", "Поставщик 3", "Цена поставщика 3"
    ]]

    # Сохранение результатов в CSV
    result_df.to_csv("matched_products.csv", index=False)
    print("Сопоставление завершено. Результаты сохранены в файл 'matched_products.csv'.")

if __name__ == "__main__":
    shop_file = r"D:\Учеба\shop_items.xlsx"
    supplier_file = r"D:\Учеба\supplier_items.xlsx"
    main(shop_file, supplier_file)
