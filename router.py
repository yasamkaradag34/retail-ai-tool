import requests
import json
import pandas as pd

OLLAMA_URL = "http://localhost:11434/api/chat"
MODEL = "llama3.1"

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "get_stock_level",
            "description": "Belirli bir ürünün stok miktarını döner",
            "parameters": {
                "type": "object",
                "properties": {
                    "product_id": {
                        "type": "string",
                        "description": "Ürün kodu, örn: SKU-001"
                    }
                },
                "required": ["product_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_daily_sales_report",
            "description": "Bugünkü satış raporunu getirir",
            "parameters": {
                "type": "object",
                "properties": {}
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_all_stock",
            "description": "Tüm ürünlerin stok listesini getirir. En fazla/az stoklu ürün, karşılaştırma gibi sorularda kullan.",
            "parameters": {
                "type": "object",
                "properties": {}
            }
        }
    }
]

def get_stock_level(product_id: str):
    df = pd.read_excel("stok.xlsx", dtype={"product_id": str})
    row = df[df["product_id"] == product_id]
    if row.empty:
        return f"{product_id} bulunamadı."
    name = row.iloc[0]["name"]
    stock = row.iloc[0]["stock"]
    return f"{name} ({product_id}) stok miktarı: {stock} adet"

def get_daily_sales_report():
    df = pd.read_excel("stok.xlsx", dtype={"product_id": str})
    return df.to_string(index=False)

def get_all_stock():
    df = pd.read_excel("stok.xlsx", dtype={"product_id": str})
    return df.to_string(index=False)

def route(user_message: str, history: list):
    history.append({"role": "user", "content": user_message})

    payload = {
        "model": MODEL,
        "messages": history,
        "tools": TOOLS,
        "stream": False
    }

    response = requests.post(OLLAMA_URL, json=payload)
    result = response.json()
    message = result.get("message", {})

    if not message:
        print("❌ Yanıt alınamadı:", result)
        return history

    tool_calls = message.get("tool_calls", [])

    if tool_calls:
        history.append(message)

        for tool_call in tool_calls:
            fn = tool_call.get("function", {})
            fn_name = fn.get("name")
            fn_args = fn.get("arguments", {})

            if fn_name == "get_stock_level":
                fn_result = get_stock_level(fn_args["product_id"])
            elif fn_name == "get_daily_sales_report":
                fn_result = get_daily_sales_report()
            elif fn_name == "get_all_stock":
                fn_result = get_all_stock()
            else:
                fn_result = "Bilinmeyen fonksiyon"

            history.append({
                "role": "tool",
                "content": fn_result
            })

        payload2 = {
            "model": MODEL,
            "messages": history,
            "stream": False
        }
        response2 = requests.post(OLLAMA_URL, json=payload2)
        result2 = response2.json()
        final_message = result2.get("message", {})
        final_content = final_message.get("content", "")

        print(f"\n🤖 {final_content}\n")
        history.append({"role": "assistant", "content": final_content})

    else:
        content = message.get("content", "")
        print(f"\n🤖 {content}\n")
        history.append({"role": "assistant", "content": content})

    return history

if __name__ == "__main__":
    print("🛒 Retail AI'ya hoş geldiniz! Çıkmak için 'q' yazın.\n")

    history = [
        {
            "role": "system",
            "content": (
                "Sen bir perakende mağazasının AI asistanısın. "
                "Stok veya satış soruları için MUTLAKA uygun tool'u çağır, asla kendin uydurmа. "
                "Tek ürün sorulursa get_stock_level, karşılaştırma/en fazla/en az için get_all_stock, "
                "satış raporu için get_daily_sales_report kullan. "
                "Türkçe konuş."
            )
        }
    ]

    while True:
        user_input = input("Sen: ").strip()
        if user_input.lower() in ["q", "quit", "çıkış"]:
            print("Görüşürüz!")
            break
        if not user_input:
            continue
        history = route(user_input, history)