from collections import defaultdict
import math

def calculate_warehouse_needs(stocks_data, default_turnover_days=30, custom_turnovers=None, round_to_5=False, use_missing_days=False):
    """
    Рассчитывает потребность к отгрузке для одного склада.
    stocks_data: список словарей остатков (с ключами 'supplierArticle', 'quantity', 'warehouseName', 'missing_days').
                 Каждая строка также содержит 'sales_period' (заказы за период) из того же Excel-отчёта.
    default_turnover_days: оборачиваемость по умолчанию для склада
    custom_turnovers: словарь {supplierArticle: days} для индивидуальной настройки
    round_to_5: Если True, потребность округляется до ближайшего числа, кратного 5, в большую сторону.
    use_missing_days: Если True, вычитает время отсутствия из периода при расчете средних дневных продаж.
    """
    if custom_turnovers is None:
        custom_turnovers = {}

    # 1. Группируем текущие остатки по артикулу продавца
    current_stocks = defaultdict(int)
    for stock in stocks_data:
        art = stock.get("supplierArticle")
        if art:
            current_stocks[str(art)] += stock.get("quantity", 0)

    # 2. Группируем продажи по артикулу продавца из загруженной таблицы
    total_sales = defaultdict(int)
    items_info = {}

    sales_period_days = 30
    if stocks_data and "report_days" in stocks_data[0]:
        sales_period_days = max(1, stocks_data[0]["report_days"])

    for stock in stocks_data:
        art = stock.get("supplierArticle")
        if art:
            art_str = str(art)
            if art_str not in items_info:
                items_info[art_str] = {
                    "nmId": stock.get("nmId", ""),
                    "subject": stock.get("itemSubject", stock.get("subject", "")),
                    "itemName": stock.get("itemName", ""),
                    "wb_turnover": stock.get("wb_turnover", 0),
                    "brand": stock.get("brand", "")
                }
            else:
                items_info[art_str]["subject"] = stock.get("itemSubject", items_info[art_str].get("subject", ""))
                items_info[art_str]["itemName"] = stock.get("itemName", items_info[art_str].get("itemName", ""))
                items_info[art_str]["wb_turnover"] = stock.get("wb_turnover", items_info[art_str].get("wb_turnover", 0))

            if "sales_period" in stock:
                total_sales[art_str] += int(stock["sales_period"])

    # 3. Рассчитываем потребность
    needs = []

    for art, info in items_info.items():
        sales_30 = float(total_sales.get(art, 0))

        missing_days = 0.0
        for stock in stocks_data:
            if str(stock.get("supplierArticle")) == str(art):
                missing_days = float(stock.get("missing_days", 0.0))
                break

        effective_days = float(sales_period_days)
        if use_missing_days:
            effective_days = max(1.0, float(sales_period_days) - missing_days)

        avg_daily_sales = float(sales_30 / effective_days)
        current_stock = int(current_stocks.get(art, 0))

        turnover_days = float(custom_turnovers.get(art) or default_turnover_days)

        # Формула: (Средние продажи в день * Целевая оборачиваемость) - Текущий остаток
        target_stock = avg_daily_sales * turnover_days

        # Если товара не было на складе весь период — фиксированная цель 100
        if use_missing_days and missing_days >= sales_period_days:
            target_stock = 100

        need = math.ceil(target_stock - float(current_stock))
        need = max(0, int(need)) # Не уводим в минус

        if round_to_5 and need > 0:
            remainder = need % 5
            if remainder > 0:
                need += (5 - remainder)

        if avg_daily_sales > 0:
            calc_wb_turnover = round(current_stock / avg_daily_sales)
        else:
            calc_wb_turnover = 0

        needs.append({
            "supplierArticle": art,
            "itemSubject": info.get("subject", ""),
            "itemName": info.get("itemName", ""),
            "wb_turnover": calc_wb_turnover,
            "nmId": info.get("nmId", ""),
            "missing_days": round(missing_days, 1),
            "avg_daily_sales": round(avg_daily_sales, 2),
            "sales_30days": sales_30,
            "current_stock": current_stock,
            "target_stock": math.ceil(target_stock),
            "turnover_days": turnover_days,
            "need": need,
            "final_shipment": need
        })

    # Возвращаем отсортированный список (сначала те, где потребность больше)
    return sorted(needs, key=lambda x: x["need"], reverse=True)
