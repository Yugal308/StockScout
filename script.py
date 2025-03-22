import pandas as pd
import yfinance as yf

def calculate_recommendation(data):
    """
    Simple recommendation system based on financial metrics
    Returns: "Buy", "Sell", or "Hold"
    """
    if not data:
        return "N/A"
    
    # Initialize score
    score = 0
    
    # Positive factors
    if data["PE Ratio"] != "N/A" and data["PE Ratio"] < 15:
        score += 1
    if data["PEG Ratio"] != "N/A" and data["PEG Ratio"] < 1:
        score += 1
    if data["Dividend Yield"] != "N/A" and data["Dividend Yield"] > 0.03:
        score += 1
    if data["Debt to Equity"] != "N/A" and data["Debt to Equity"] < 1:
        score += 1
    
    # Negative factors
    if data["PE Ratio"] != "N/A" and data["PE Ratio"] > 25:
        score -= 1
    if data["PEG Ratio"] != "N/A" and data["PEG Ratio"] > 1.5:
        score -= 1
    if data["Debt to Equity"] != "N/A" and data["Debt to Equity"] > 2:
        score -= 1
    if data["Current Price"] != "N/A" and data["52 Week High"] != "N/A" and \
       data["Current Price"] > 0.9 * data["52 Week High"]:
        score -= 1
    
    # Determine recommendation
    if score >= 2:
        return "Buy"
    elif score <= -2:
        return "Sell"
    return "Hold"

def fetch_stock_data(symbol, exchange):
    try:
        # Append the correct suffix based on the exchange
        if exchange == "NSE":
            stock_symbol = f"{symbol}.NS"
        elif exchange == "BSE":
            stock_symbol = f"{symbol}.BO"
        else:
            return None
        
        stock = yf.Ticker(stock_symbol)
        hist = stock.history(period="max")
        
        if hist.empty:
            return None
        
        return {
            "Current Price": stock.info.get("currentPrice", "N/A"),
            "Previous Close": stock.info.get("previousClose", "N/A"),
            "Open Price": stock.info.get("open", "N/A"),
            "Day High": stock.info.get("dayHigh", "N/A"),
            "Day Low": stock.info.get("dayLow", "N/A"),
            "All Time High": hist["High"].max(),
            "All Time Low": hist["Low"].min(),
            "52 Week High": stock.info.get("fiftyTwoWeekHigh", "N/A"),
            "52 Week Low": stock.info.get("fiftyTwoWeekLow", "N/A"),
            "PE Ratio": stock.info.get("trailingPE", "N/A"),
            "PEG Ratio": stock.info.get("pegRatio", "N/A"),
            "Price to Sales": stock.info.get("priceToSalesTrailing12Months", "N/A"),
            "Price to Book": stock.info.get("priceToBook", "N/A"),
            "Market Cap": stock.info.get("marketCap", "N/A"),
            "Total Revenue": stock.info.get("totalRevenue", "N/A"),
            "Gross Profits": stock.info.get("grossProfits", "N/A"),
            "EBITDA": stock.info.get("ebitda", "N/A"),
            "Total Cash": stock.info.get("totalCash", "N/A"),
            "Total Debt": stock.info.get("totalDebt", "N/A"),
            "Debt to Equity": stock.info.get("debtToEquity", "N/A"),
            "Current Ratio": stock.info.get("currentRatio", "N/A"),
            "Quick Ratio": stock.info.get("quickRatio", "N/A"),
            "Dividend Yield": stock.info.get("dividendYield", "N/A"),
            "Payout Ratio": stock.info.get("payoutRatio", "N/A"),
            "Volume": stock.info.get("volume", "N/A"),
            "Beta": stock.info.get("beta", "N/A"),
            "Sector": stock.info.get("sector", "N/A"),
            "Industry": stock.info.get("industry", "N/A"),
            "Insider Holdings %": stock.info.get("heldPercentInsiders", "N/A"),
            "Institutional Holdings %": stock.info.get("heldPercentInstitutions", "N/A")
        }
    except Exception as e:
        print(f"Error fetching data for {symbol}: {e}")
        return None

def process_excel(file_path):
    df = pd.read_excel(file_path)
    
    new_columns = [
        "Current Price", "Previous Close", "Open Price", "Day High", "Day Low", "All Time High", 
        "All Time Low", "52 Week High", "52 Week Low", "PE Ratio", "PEG Ratio", "Price to Sales", 
        "Price to Book", "Market Cap", "Total Revenue", "Gross Profits", "EBITDA", "Total Cash", 
        "Total Debt", "Debt to Equity", "Current Ratio", "Quick Ratio", "Dividend Yield", 
        "Payout Ratio", "Volume", "Beta", "Sector", "Industry", "Insider Holdings %", 
        "Institutional Holdings %", "Recommendation"
    ]
    
    for col in new_columns:
        df[col] = None
    
    for index, row in df.iterrows():
        stock_symbol = row["SYMBOL"].strip()
        exchange = row.get("Exchange", "NSE")
        data = fetch_stock_data(stock_symbol, exchange)
        
        if data:
            for col in new_columns[:-1]:
                df.at[index, col] = data[col]
            df.at[index, "Recommendation"] = calculate_recommendation(data)
    
    # Save to a temporary file
    output_file = "updated_stock_data.xlsx"
    df.to_excel(output_file, index=False)
    return output_file
