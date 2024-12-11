import requests
import pandas as pd
import schedule
import time

def fetch_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    response = requests.get(url, params=params)
    data = response.json()
    return data


def analyze_data(data):
    
    df = pd.DataFrame(data)
    
    
    print("Data Overview:")
    print(df.head(), "\n")

    
    top_5_market_cap = df[['name', 'symbol', 'market_cap']].sort_values(by='market_cap', ascending=False).head(5)
    print("Top 5 Cryptocurrencies by Market Cap:")
    print(top_5_market_cap, "\n")

    
    avg_price = df['current_price'].mean()
    print(f"Average Price of Top 50 Cryptocurrencies: ${avg_price:.2f}\n")

    highest_change = df[['name', 'symbol', 'price_change_percentage_24h']].sort_values(by='price_change_percentage_24h', ascending=False).head(1)
    lowest_change = df[['name', 'symbol', 'price_change_percentage_24h']].sort_values(by='price_change_percentage_24h').head(1)
    
    print("Cryptocurrency with Highest 24-Hour Price Change:")
    print(highest_change, "\n")
    
    print("Cryptocurrency with Lowest 24-Hour Price Change:")
    print(lowest_change, "\n")

    analysis_results = pd.DataFrame({
        'Top 5 Market Cap': top_5_market_cap['name'].values,
        'Market Cap (USD)': top_5_market_cap['market_cap'].values,
        'Average Price': [avg_price] * len(top_5_market_cap),
        'Highest 24h Change': [highest_change['name'].values[0]] * len(top_5_market_cap),
        'Lowest 24h Change': [lowest_change['name'].values[0]] * len(top_5_market_cap)
    })


    save_analysis_to_excel(df, analysis_results)

def save_analysis_to_excel(df, analysis_results):
    
    with pd.ExcelWriter('crypto_analysis.xlsx', engine='openpyxl') as writer:
        
        df.to_excel(writer, sheet_name='Live Data', index=False)
        
        analysis_results.to_excel(writer, sheet_name='Analysis', index=False)
    
    print("Data and analysis have been successfully saved to 'crypto_analysis.xlsx'")










if __name__ == "__main__":
    data = fetch_data()
    analyze_data(data)
    
    
    
def job():
    data = fetch_data()
    analyze_data(data)
    

schedule.every(5).minutes.do(job)  # Run every 5 minutes

while True:
    schedule.run_pending()
    time.sleep(1)    
