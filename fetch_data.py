import pandas as pd                                                     # Import pandas for data handling and DataFrame manipulation
import yfinance as yf                                                   # Import yfinance to fetch stock market data
import os                                                               # Import os for file and path handling
import sys                                                              # Import sys to modify Python path for module imports


# ------✅ Add the scripts directory to Python path------

scripts_dir = os.path.dirname(os.path.abspath(__file__))                # Get the directory where this script is located
if scripts_dir not in sys.path:                                         # Check if that directory is not already in Python's search path
    sys.path.append(scripts_dir)                                        # Add it to Python’s module search path


# ------✅ Import indicators with fallback------

try:
    from indicators import calculate_rsi, calculate_macd                # Try importing custom RSI and MACD functions
    print("✅ Imported indicators from local module")
except ImportError:
    try:
        from .indicators import calculate_rsi, calculate_macd           # Alternative import style (relative import if using packages)
        print("✅ Imported indicators from relative module")
    except ImportError:
        print("❌ Could not import indicators module")
        # Define fallback functions if module not found
        def calculate_rsi(close_prices, period=14):
            delta = close_prices.diff()                                 # Price change from previous day
            gain = delta.where(delta > 0, 0)                            # Positive changes only
            loss = -delta.where(delta < 0, 0)                           # Negative changes only
            avg_gain = gain.rolling(window=period).mean()               # Average gain over 'period'
            avg_loss = loss.rolling(window=period).mean()               # Average loss over 'period'
            rs = avg_gain / avg_loss                                    # Relative strength (RS)
            rsi = 100 - (100 / (1 + rs))                                # Final RSI formula
            return rsi

        def calculate_macd(close_prices, short_window=12, long_window=26, signal_window=9):
            ema_short = close_prices.ewm(span=short_window, adjust=False).mean()    # 12-day EMA
            ema_long = close_prices.ewm(span=long_window, adjust=False).mean()      # 26-day EMA
            macd = ema_short - ema_long                                             # MACD = short EMA - long EMA
            signal = macd.ewm(span=signal_window, adjust=False).mean()              # Signal line (9-day EMA of MACD)
            histogram = macd - signal                                               # Difference = histogram
            return macd, signal, histogram


# ------✅ Read tickers from Excel file------
def get_tickers_from_excel(excel_path=None, sheet_name="Sheet1"):
    if excel_path is None:                                                          # If no path is provided, locate dashboard.xlsm automatically
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))      # Go one folder up
        excel_path = os.path.join(base_dir, "dashboard.xlsm")                       # Default Excel file name

    print(f"📁 Looking for Excel file at: {excel_path}")

    if not os.path.exists(excel_path):                                              # If the file does not exist
        print(f"❌ Excel file not found at: {excel_path}")
        return ["AAPL", "MSFT", "GOOGL", "TSLA", "AMZN"]                            # Return fallback tickers

    try:
        excel_file = pd.ExcelFile(excel_path)                                       # Load Excel file to check sheet names
        print(f"✅ Available sheets: {excel_file.sheet_names}")

        df = pd.read_excel(excel_path, sheet_name=sheet_name)                       # Read the specified sheet
        print(f"✅ Successfully read sheet: {sheet_name}")
        print(f"Columns found: {df.columns.tolist()}")                              # Show columns found in Excel

        tickers = df.iloc[:, 0].dropna().tolist()                                   # Read tickers from first column
        print(f"✅ Tickers extracted: {tickers}")
        return tickers

    except Exception as e:                                                          # If any error occurs
        print(f"❌ Error reading Excel file: {e}")
        return ["AAPL", "MSFT", "GOOGL", "TSLA", "AMZN"]                            # Return fallback tickers


# ------✅ Fetch data and calculate indicators for each ticker------
def fetch_stock_data_with_indicators(tickers):
    all_data = []                                                                   # List to store data of all stocks

    for ticker in tickers:                                                          # Loop through each ticker
        try:
            print(f"📊 Fetching data for {ticker}...")
            stock = yf.Ticker(ticker)                                               # Create a ticker object

            info = stock.info                                                       # Fetch fundamental info
            hist = stock.history(period="6mo")                                      # Fetch last 6 months historical data

            # ✅ Extract key metrics
            current_price = info.get("currentPrice") or info.get("regularMarketPrice")
            pe_ratio = info.get("trailingPE") or info.get("forwardPE")
            market_cap = info.get("marketCap")
            dividend_yield = info.get("dividendYield")

            # ✅ Calculate RSI & MACD only if history is valid
            if not hist.empty and len(hist) > 14:
                close_prices = hist["Close"]
                rsi = calculate_rsi(close_prices).iloc[-1]                          # Latest RSI value
                macd, signal, _ = calculate_macd(close_prices)                      # Full MACD calculation
                macd_value = macd.iloc[-1]
                signal_value = signal.iloc[-1]
            else:
                rsi = macd_value = signal_value = None

            data = {  # Prepare data row for this stock
                "Ticker": ticker,
                "Current Price": round(current_price, 2) if current_price else "N/A",
                "PE Ratio": round(pe_ratio, 2) if pe_ratio else "N/A",
                "Market Cap": f"{round(market_cap / 1e9, 2)}B" if market_cap else "N/A",
                "Dividend Yield": round(dividend_yield, 4) if dividend_yield else "N/A",
                "RSI (14)": round(rsi, 2) if rsi else "N/A",
                "MACD": round(macd_value, 2) if macd_value is not None else "N/A",
                "Signal Line": round(signal_value, 2) if signal_value is not None else "N/A",
            }

            all_data.append(data)                                                   # Add to list
            print(f"✅ Successfully processed {ticker}")

        except Exception as e:                                                      # If stock data fetching fails
            print(f"❌ Error fetching {ticker}: {e}")
            all_data.append({
                "Ticker": ticker, "Current Price": "Error", "PE Ratio": "Error",
                "Market Cap": "Error", "Dividend Yield": "Error", "RSI (14)": "Error",
                "MACD": "Error", "Signal Line": "Error"
            })

    return pd.DataFrame(all_data)                                                   # Convert list to DataFrame

def main():
    tickers = get_tickers_from_excel()                                              # Step 1: Read stock tickers
    print("✅ Tickers found:", tickers)

    df = fetch_stock_data_with_indicators(tickers)                                  # Step 2: Fetch stock data & indicators 
    print(df)

    # # Save to Excel for verification
    # output_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "stock_data_output.xlsx")
    # df.to_excel(output_path, index=False)
    # print(f"📁 Saved to {output_path}")

if __name__ == "__main__":                                                          # Run only if file is executed directly
    main()