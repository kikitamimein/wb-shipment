import requests
import urllib.parse

MS_API_TOKEN = ""
MS_WB_ACCOUNT = "WB1"
MS_API_URL = "https://api.moysklad.ru/api/remap/1.2"

def get_headers():
    return {
        "Authorization": f"Bearer {MS_API_TOKEN}",
        "Accept": "application/json;charset=utf-8"
    }

def get_product(barcode_or_article):
    """Ищет товар в МС по штрихкоду или артикулу."""
    # сначала попробуем по баркоду, если не нашло, то по артикулу
    url = f"{MS_API_URL}/entity/product"
    
    for search_term in [barcode_or_article]:
        # Simple search using the search param. We can also use filter=barcode=...
        params = {"search": str(search_term)}
        response = requests.get(url, headers=get_headers(), params=params)
        response.raise_for_status()
        data = response.json()
        
        if data.get('rows') and len(data['rows']) > 0:
            return data['rows'][0]
    
    return None

def get_organization(name="KeshFix"):
    url = f"{MS_API_URL}/entity/organization"
    response = requests.get(url, headers=get_headers(), params={"search": name})
    response.raise_for_status()
    rows = response.json().get('rows', [])
    for r in rows:
        if r['name'] == name:
            return r['meta']
    if rows:
        return rows[0]['meta']
    return None

def get_store(name="основной"):
    url = f"{MS_API_URL}/entity/store"
    response = requests.get(url, headers=get_headers(), params={"search": name})
    response.raise_for_status()
    rows = response.json().get('rows', [])
    for r in rows:
        if r['name'] == name:
            return r['meta']
    if rows:
        return rows[0]['meta']
    return None

def get_counterparty_meta(cp_id="126a51cf-5d43-11ee-0a80-11600012c549"):
    return {
        "href": f"{MS_API_URL}/entity/counterparty/{cp_id}",
        "type": "counterparty",
        "mediaType": "application/json"
    }

def get_state(entity_type="demand", name="Создаётся поставка"):
    url = f"{MS_API_URL}/entity/{entity_type}/metadata"
    response = requests.get(url, headers=get_headers())
    response.raise_for_status()
    states = response.json().get('states', [])
    for s in states:
        if s['name'].lower() == name.lower():
            return s['meta']
    return None

def create_demand_v2(positions_data, org_meta, store_meta, cp_meta, state_meta=None, name="", description="", shipment_address=""):
    """
    positions_data: list of dicts {"meta": product_meta, "quantity": int, "price": int} (price in cents)
    """
    url = f"{MS_API_URL}/entity/demand"
    
    payload = {
        "organization": {"meta": org_meta},
        "agent": {"meta": cp_meta},
        "store": {"meta": store_meta},
    }
    
    if name:
        payload["name"] = name
    if description:
        payload["description"] = description
    if state_meta:
        payload["state"] = {"meta": state_meta}
    if shipment_address:
        payload["shipmentAddress"] = shipment_address
        payload["description"] = description
    if state_meta:
        payload["state"] = {"meta": state_meta}
        
    lines = []
    for p in positions_data:
        try:
            qty = float(p["quantity"])
        except:
            qty = 0
            
        line = {
            "quantity": qty,
            "assortment": {"meta": p["meta"]}
        }
        if "price" in p and p["price"] is not None:
            line["price"] = int(p["price"])
        lines.append(line)
        
    payload["positions"] = lines
    
    response = requests.post(url, headers=get_headers(), json=payload)
    if response.status_code not in [200, 201]:
        raise Exception(f"Ошибка создания отгрузки: {response.text}")
    return response.json()

def get_demands(days_back=14, limit=200):
    """Возвращает список отгрузок (demands) за последние days_back дней."""
    from datetime import datetime, timedelta
    cutoff = (datetime.now() - timedelta(days=days_back)).strftime("%Y-%m-%d %H:%M:%S")

    url = f"{MS_API_URL}/entity/demand"
    params = {
        "limit": min(limit, 100),
        "offset": 0,
        "order": "moment,desc",
        "filter": f"moment>={cutoff}"
    }

    demands = []
    while True:
        resp = requests.get(url, headers=get_headers(), params=params)
        if resp.status_code != 200:
            raise Exception(f"Ошибка получения отгрузок: {resp.text}")
        data = resp.json()
        rows = data.get("rows", [])
        for r in rows:
            demands.append({
                "name": r.get("name", ""),
                "id": r.get("id", ""),
                "meta": r.get("meta"),
                "shipmentAddress": r.get("shipmentAddress", ""),
                "state": r.get("state", {}).get("name", ""),
                "moment": r.get("moment", ""),
                "description": r.get("description", ""),
            })
        if len(rows) < 100 or len(demands) >= limit:
            break
        params["offset"] += 100
    return demands[:limit]


def get_demand_positions(demand_href):
    """Возвращает список позиций отгрузки. Каждая позиция содержит assortment meta, quantity, id."""
    url = f"{demand_href}/positions"
    params = {"limit": 1000, "offset": 0}
    positions = []
    while True:
        resp = requests.get(url, headers=get_headers(), params=params)
        if resp.status_code != 200:
            raise Exception(f"Ошибка получения позиций: {resp.text}")
        data = resp.json()
        rows = data.get("rows", [])
        for r in rows:
            positions.append({
                "id": r.get("id"),
                "quantity": float(r.get("quantity", 0)),
                "price": float(r.get("price", 0)) / 100,
                "meta": r.get("assortment", {}).get("meta"),
            })
        if len(rows) < 1000:
            break
        params["offset"] += 1000
    return positions


def update_position(demand_href, position_id, quantity, price=None):
    """Обновляет количество (и опционально цену) позиции в отгрузке."""
    url = f"{demand_href}/positions/{position_id}"
    payload = {"quantity": float(quantity)}
    if price is not None:
        payload["price"] = int(price * 100)
    resp = requests.put(url, headers=get_headers(), json=payload)
    if resp.status_code not in [200, 201]:
        raise Exception(f"Ошибка обновления позиции: {resp.text}")
    return resp.json()


def delete_position(demand_href, position_id):
    """Удаляет позицию из отгрузки."""
    url = f"{demand_href}/positions/{position_id}"
    resp = requests.delete(url, headers=get_headers())
    if resp.status_code not in [200, 204]:
        raise Exception(f"Ошибка удаления позиции: {resp.text}")


def get_recent_supplies(days_back=14):
    """Возвращает словарь {артикул: дата_последней_приёмки} для товаров,
    принятых за последние days_back дней."""
    from datetime import datetime, timedelta
    cutoff = (datetime.now() - timedelta(days=days_back)).strftime("%Y-%m-%d %H:%M:%S")

    url = f"{MS_API_URL}/entity/supply"
    params = {
        "limit": 100,
        "offset": 0,
        "order": "moment,desc",
        "expand": "positions.assortment",
        "filter": f"moment>={cutoff}"
    }

    recent = {}
    while True:
        resp = requests.get(url, headers=get_headers(), params=params)
        if resp.status_code != 200:
            raise Exception(f"Ошибка получения приёмок: {resp.text}")
        data = resp.json()
        rows = data.get("rows", [])

        for supply in rows:
            moment = supply.get("moment", "")
            positions = supply.get("positions", {})
            pos_rows = positions.get("rows", []) if isinstance(positions, dict) else []

            for pos in pos_rows:
                assortment = pos.get("assortment", {})
                art = assortment.get("article") if isinstance(assortment, dict) else None
                if art:
                    art = str(art)
                    if art not in recent or moment > recent[art]:
                        recent[art] = moment

        if len(rows) < 100:
            break
        params["offset"] += 100

    return recent


def get_recent_supplies_detailed(days_back=60):
    """Возвращает список словарей с детальной информацией о последних приёмках.

    Каждый словарь: article, name, barcode, received_date, quantity, meta (assortment).
    По одному товару — самая поздняя приёмка за период.
    """
    from datetime import datetime, timedelta
    cutoff = (datetime.now() - timedelta(days=days_back)).strftime("%Y-%m-%d %H:%M:%S")

    url = f"{MS_API_URL}/entity/supply"
    params = {
        "limit": 100,
        "offset": 0,
        "order": "moment,desc",
        "expand": "positions.assortment",
        "filter": f"moment>={cutoff}"
    }

    seen = {}
    while True:
        resp = requests.get(url, headers=get_headers(), params=params)
        if resp.status_code != 200:
            raise Exception(f"Ошибка получения приёмок: {resp.text}")
        data = resp.json()
        rows = data.get("rows", [])

        for supply in rows:
            moment = supply.get("moment", "")
            positions = supply.get("positions", {})
            pos_rows = positions.get("rows", []) if isinstance(positions, dict) else []

            for pos in pos_rows:
                assortment = pos.get("assortment", {})
                if not isinstance(assortment, dict):
                    continue
                art = assortment.get("article")
                if not art:
                    continue
                art = str(art)

                if art in seen:
                    continue

                barcodes = assortment.get("barcodes", [])
                barcode = ""
                if barcodes:
                    b = barcodes[0]
                    barcode = b.get("code128") or b.get("ean13") or ""

                seen[art] = {
                    "article": art,
                    "name": assortment.get("name", ""),
                    "barcode": str(barcode) if barcode else "",
                    "received_date": moment,
                    "quantity": float(pos.get("quantity", 0)),
                    "meta": assortment.get("meta"),
                }

        if len(rows) < 100:
            break
        params["offset"] += 100

    return list(seen.values())


def add_positions_to_demand(demand_href, positions):
    """Добавляет позиции к существующей отгрузке.

    positions: list of dicts {"meta": product_meta, "quantity": int, "price": int} (price in cents)
    """
    url = f"{demand_href}/positions"
    lines = []
    for p in positions:
        qty = float(p.get("quantity", 0))
        line = {"quantity": qty, "assortment": {"meta": p["meta"]}}
        if "price" in p and p["price"] is not None:
            line["price"] = int(p["price"])
        lines.append(line)

    resp = requests.post(url, headers=get_headers(), json=lines)
    if resp.status_code not in [200, 201]:
        raise Exception(f"Ошибка добавления позиций: {resp.text}")
    return resp.json()


def get_all_stocks():
    """Возвращает словарь {артикул: остаток} для всех товаров, где остаток = физический запас на складах."""
    url = f"{MS_API_URL}/report/stock/all"
    stocks = {}
    
    params = {"limit": 1000, "offset": 0}
    while True:
        response = requests.get(url, headers=get_headers(), params=params)
        if response.status_code != 200:
            print(f"Error fetching MS stocks: {response.text}")
            break
            
        data = response.json()
        rows = data.get('rows', [])
        
        for row in rows:
            article = row.get("article")
            if article:
                stocks[str(article)] = {
                    "stock": float(row.get("stock", 0)),
                    "price": float(row.get("salePrice", 0)),
                    "meta": row.get("meta"),
                    "folder": row.get("folder", {}).get("name", "Без группы")
                }
                
        if len(rows) < 1000:
            break
        params["offset"] += 1000
        
    return stocks

