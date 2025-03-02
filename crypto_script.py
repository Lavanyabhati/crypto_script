import requests
import pandas as pd
import time
from openpyxl import Workbook

API_URL = "https://api.coingecko.com/api/v3/coins/markets"
PARAMS = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": False
}

def fetch_crypto_data():
    response = requests.get(API_URL, params=PARAMS)
    if response.status_code == 200:
        return response.json()
    else:
        print("Error fetching data")
        return []

def analyze_data(data):
    df = pd.DataFrame(data, columns=["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"])

    top_5 = df.nlargest(5, "market_cap")

    avg_price = df["current_price"].mean()
    
    highest_change = df.loc[df["price_change_percentage_24h"].idxmax()]
    lowest_change = df.loc[df["price_change_percentage_24h"].idxmin()]
    
    return top_5, avg_price, highest_change, lowest_change, df

def update_excel():
    print("Starting Excel update...")
    wb = Workbook()
    ws = wb.active
    ws.title = "Crypto Data"
    
    while True:
        data = fetch_crypto_data()
        if not data:
            continue
        
        top_5, avg_price, highest_change, lowest_change, df = analyze_data(data)

        ws.append(["Name", "Symbol", "Price (USD)", "Market Cap", "Volume", "24h Change (%)"])
        for _, row in df.iterrows():
            ws.append(row.tolist())

        wb.save("crypto_data.xlsx")
        print("Excel updated. Next update in 5 minutes...")
        time.sleep(300)  

update_excel()

