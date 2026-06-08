import pandas as pd
from datetime import datetime

ORDERS_PATH = "orders.xlsx"

def get_order_status(order_id: str):
    df = pd.read_excel(ORDERS_PATH, dtype={"order_id": str})
    row = df[df["order_id"] == order_id]
    if row.empty:
        return f"{order_id} numaralı sipariş bulunamadı."
    status = row.iloc[0]["status"]
    customer = row.iloc[0]["customer"]
    return f"Sipariş {order_id} — Müşteri: {customer}, Durum: {status}"

def get_all_orders():
    df = pd.read_excel(ORDERS_PATH, dtype={"order_id": str})
    return df.to_string(index=False)

def get_pending_orders():
    df = pd.read_excel(ORDERS_PATH, dtype={"order_id": str})
    pending = df[df["status"] == "Beklemede"]
    if pending.empty:
        return "Bekleyen sipariş yok."
    return pending.to_string(index=False)

def get_orders_by_customer(customer: str):
    df = pd.read_excel(ORDERS_PATH, dtype={"order_id": str})
    result = df[df["customer"].str.contains(customer, case=False, na=False)]
    if result.empty:
        return f"'{customer}' adlı müşteriye ait sipariş bulunamadı."
    return result.to_string(index=False)

def update_order_status(order_id: str, status: str):
    df = pd.read_excel(ORDERS_PATH, dtype={"order_id": str})
    if order_id not in df["order_id"].values:
        return f"{order_id} numaralı sipariş bulunamadı."
    df.loc[df["order_id"] == order_id, "status"] = status
    df.to_excel(ORDERS_PATH, index=False)
    return f"Sipariş {order_id} durumu '{status}' olarak güncellendi."

def get_todays_orders():
    df = pd.read_excel(ORDERS_PATH, dtype={"order_id": str})
    today = datetime.today().strftime("%Y-%m-%d")
    df["date"] = pd.to_datetime(df["date"]).dt.strftime("%Y-%m-%d")
    todays = df[df["date"] == today]
    if todays.empty:
        return "Bugün henüz sipariş yok."
    return todays.to_string(index=False)