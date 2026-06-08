import pandas as pd

EXCEL_PATH = "stok.xlsx"

def get_stock_level(product_id: str):
    df = pd.read_excel(EXCEL_PATH, dtype={"product_id": str})
    row = df[df["product_id"] == product_id]
    if row.empty:
        return f"{product_id} bulunamadı."
    name = row.iloc[0]["name"]
    stock = row.iloc[0]["stock"]
    return f"{name} ({product_id}) stok miktarı: {stock} adet"

def get_all_stock():
    df = pd.read_excel(EXCEL_PATH, dtype={"product_id": str})
    return df.to_string(index=False)

def get_daily_sales_report():
    df = pd.read_excel(EXCEL_PATH, dtype={"product_id": str})
    return df.to_string(index=False)

def check_low_stock(threshold: int = 10):
    df = pd.read_excel(EXCEL_PATH, dtype={"product_id": str})
    low = df[df["stock"] <= threshold]
    if low.empty:
        return f"Stoku {threshold} adetten az olan ürün yok."
    return low.to_string(index=False)

def get_out_of_stock():
    df = pd.read_excel(EXCEL_PATH, dtype={"product_id": str})
    out = df[df["stock"] == 0]
    if out.empty:
        return "Tükenmiş ürün yok."
    return out.to_string(index=False)

def search_product_by_name(name: str):
    df = pd.read_excel(EXCEL_PATH, dtype={"product_id": str})
    result = df[df["name"].str.contains(name, case=False, na=False)]
    if result.empty:
        return f"'{name}' adında ürün bulunamadı."
    return result.to_string(index=False)

def get_stock_value():
    df = pd.read_excel(EXCEL_PATH, dtype={"product_id": str})
    df["total_value"] = df["stock"] * df["price"]
    total = df["total_value"].sum()
    detail = df[["product_id", "name", "stock", "price", "total_value"]].to_string(index=False)
    return f"{detail}\n\nToplam Stok Değeri: {total:.2f} TL"

def update_stock(product_id: str, quantity: int):
    df = pd.read_excel(EXCEL_PATH, dtype={"product_id": str})
    if product_id not in df["product_id"].values:
        return f"{product_id} bulunamadı."
    df.loc[df["product_id"] == product_id, "stock"] = quantity
    df.to_excel(EXCEL_PATH, index=False)
    name = df[df["product_id"] == product_id].iloc[0]["name"]
    return f"{name} ({product_id}) stok miktarı {quantity} olarak güncellendi."