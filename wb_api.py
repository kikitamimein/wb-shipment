import datetime
import pandas as pd
import os
import re

def get_stocks_from_excel(filepath="остатки.xlsx"):
    """Читает остатки из скачанного Excel-файла WB (4 лист)."""
    if not os.path.exists(filepath):
        raise Exception(f"Файл {filepath} не найден. Пожалуйста, выберите файл в интерфейсе.")

    xl = pd.ExcelFile(filepath)
    if len(xl.sheet_names) < 4:
        raise Exception("В файле меньше 4 листов. Убедитесь, что это корректный отчет с WB.")

    report_days = 30
    try:
        info_df = xl.parse(xl.sheet_names[0])
        for _, r in info_df.iterrows():
            if str(r.iloc[0]).lower().strip() == "выбранный период":
                val_str = str(r.iloc[1])
                matches = re.findall(r"(\d{4}-\d{2}-\d{2})", val_str)
                if len(matches) >= 2:
                    d1 = datetime.datetime.strptime(matches[0], "%Y-%m-%d")
                    d2 = datetime.datetime.strptime(matches[1], "%Y-%m-%d")
                    report_days = (d2 - d1).days + 1
                break
    except Exception:
        pass

    sheet_name = xl.sheet_names[3]
    df = xl.parse(sheet_name, skiprows=1) # Пропускаем первую строку с надписью

    # Пытаемся найти нужные колонки
    article_col = next((c for c in df.columns if "артикул продавца" in str(c).lower() or "supplierarticle" in str(c).lower()), None)
    quantity_col = next((c for c in df.columns if "остатки на текущий день" in str(c).lower() or "остаток" in str(c).lower() or "количество" in str(c).lower()), None)

    if not article_col or not quantity_col:
        raise Exception(f"Не удалось распознать колонки (Артикул продавца/Остатки) в файле {filepath}. Колонки: {list(df.columns)}")


    avail_col = next((c for c in df.columns if "доступность" in str(c).lower()), None)

    stocks = []
    for _, row in df.iterrows():
        # Check if the item is obsolete/deleted from WB
        if avail_col:
            avail_val = str(row[avail_col]).strip().lower()
            if avail_val == "неактуальный":
                continue

        art = row[article_col]
        qty = row[quantity_col]

        # Склад
        wh_col = next((c for c in df.columns if "склад" in str(c).lower()), None)
        warehouse = row[wh_col] if wh_col else ""

        # Предмет / Категория
        subject_col = next((c for c in df.columns if "предмет" in str(c).lower()), None)
        item_subject = row[subject_col] if subject_col else ""

        # Название товара
        name_col = next((c for c in df.columns if "название" in str(c).lower() or "наименование" in str(c).lower()), None)
        item_name = row[name_col] if name_col else ""

        # Оборачиваемость
        turnover_col = next((c for c in df.columns if "оборачиваемость" in str(c).lower()), None)
        wb_turnover = row[turnover_col] if turnover_col else 0
        wb_turnover = 0 if pd.isna(wb_turnover) else wb_turnover

        # Заказы — ищем колонку гибко: "заказали", "заказано", "продажи", "выкупили"
        sales_col = next((c for c in df.columns if any(kw in str(c).lower() for kw in ("заказа", "продаж", "выкуп"))), None)
        sales_period = row[sales_col] if sales_col else 0
        sales_period = 0 if pd.isna(sales_period) else int(sales_period)

        out_col = next((c for c in df.columns if "время отсутствия" in str(c).lower()), None)
        out_val = row[out_col] if out_col else "0ч"


        missing_days = 0.0
        if pd.notna(out_val):
            val_str = str(out_val).lower()
            d_match = re.search(r'(\d+)\s*д', val_str)
            h_match = re.search(r'(\d+)\s*ч', val_str)
            d_num = float(d_match.group(1)) if d_match else 0.0
            h_num = float(h_match.group(1)) if h_match else 0.0
            missing_days = d_num + (h_num / 24.0)

        # Safely convert turnover to number, or keep as string if it includes text like "11ч"
        try:
            wb_turnover_safe = float(str(wb_turnover).replace(',', '.').replace(' ', ''))
        except ValueError:
            wb_turnover_safe = str(wb_turnover)

        if pd.notna(art) and pd.notna(qty):
            stocks.append({
                "supplierArticle": str(art).strip(),
                "quantity": int(qty),
                "warehouseName": str(warehouse).strip() if pd.notna(warehouse) else "Неизвестно",
                "itemSubject": str(item_subject).strip() if pd.notna(item_subject) else "",
                "itemName": str(item_name).strip() if pd.notna(item_name) else "",
                "wb_turnover": wb_turnover_safe,
                "missing_days": missing_days,
                "sales_period": sales_period,
                "report_days": report_days
            })

    return stocks
