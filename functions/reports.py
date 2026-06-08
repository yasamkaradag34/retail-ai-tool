import pandas as pd
from datetime import datetime, timedelta

ORDERS_PATH = "orders.xlsx"
STOCK_PATH = "stok.xlsx"

def get_total_revenue():
    df_orders = pd.read_excel(ORDERS_PATH, dtype={"order_id": str})
    df_stock = pd.read_excel(STOCK_PATH, dtype={"product_id": str})
    df = df_orders.merge(df_stock[["product_id", "price"]], on="product_id", how="left")
    df = df[df["status"] == "Teslim Edildi"]
    df["revenue"] = df["quantity"] * df["price"]
    total = df["revenue"].sum()
    return f"Toplam gelir (teslim edilenler): {total:.2f} TL"

def get_best_selling_product():
    df_orders = pd.read_excel(ORDERS_PATH, dtype={"order_id": str})
    df_stock = pd.read_excel(STOCK_PATH, dtype={"product_id": str})
    df = df_orders[df_orders["status"] != "İptal"]
    grouped = df.groupby("product_id")["quantity"].sum().reset_index()
    grouped = grouped.merge(df_stock[["product_id", "name"]], on="product_id", how="left")
    grouped = grouped.sort_values("quantity", ascending=False)
    if grouped.empty:
        return "Satış verisi bulunamadı."
    top = grouped.iloc[0]
    return f"En çok satan ürün: {top['name']} ({top['product_id']}) — {top['quantity']} adet"

def get_sales_summary():
    df_orders = pd.read_excel(ORDERS_PATH, dtype={"order_id": str})
    df_stock = pd.read_excel(STOCK_PATH, dtype={"product_id": str})
    total = len(df_orders)
    delivered = len(df_orders[df_orders["status"] == "Teslim Edildi"])
    pending = len(df_orders[df_orders["status"] == "Beklemede"])
    cancelled = len(df_orders[df_orders["status"] == "İptal"])
    df = df_orders.merge(df_stock[["product_id", "price"]], on="product_id", how="left")
    df = df[df["status"] == "Teslim Edildi"]
    df["revenue"] = df["quantity"] * df["price"]
    revenue = df["revenue"].sum()
    return (
        f"Sipariş Özeti:\n"
        f"Toplam Sipariş: {total}\n"
        f"Teslim Edildi: {delivered}\n"
        f"Beklemede: {pending}\n"
        f"İptal: {cancelled}\n"
        f"Toplam Gelir: {revenue:.2f} TL"
    )

def get_low_stock_report():
    df = pd.read_excel(STOCK_PATH, dtype={"product_id": str})
    low = df[df["stock"] <= 10]
    if low.empty:
        return "Kritik stok seviyesinde ürün yok."
    return f"Kritik stok uyarısı:\n{low.to_string(index=False)}"

def get_stock_turnover():
    df_orders = pd.read_excel(ORDERS_PATH, dtype={"order_id": str})
    df_stock = pd.read_excel(STOCK_PATH, dtype={"product_id": str})
    df = df_orders[df_orders["status"] != "İptal"]
    grouped = df.groupby("product_id")["quantity"].sum().reset_index()
    grouped = grouped.merge(df_stock[["product_id", "name", "stock"]], on="product_id", how="left")
    grouped.columns = ["product_id", "sold", "name", "current_stock"]
    return grouped[["product_id", "name", "sold", "current_stock"]].to_string(index=False)