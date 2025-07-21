import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta

# Helper functions
def get_price_n_days_ago(ticker_obj, days):
    try:
        end_date = datetime.today()
        start_date = end_date - timedelta(days=days)
        hist = ticker_obj.history(start=start_date, end=end_date)
        if not hist.empty:
            return hist['Close'].iloc[0]
        else:
            return None
    except:
        return None

def get_rsi(ticker):
    try:
        df = yf.download(ticker, period="6mo", interval="1d", auto_adjust=False)
        delta = df['Close'].diff()
        gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
        rs = gain / loss
        rsi = 100 - (100 / (1 + rs))
        return float(rsi.values[-1])  # Extract raw float value
    except:
        return None

def get_macd_signal(ticker):
    try:
        df = yf.download(ticker, period="6mo", interval="1d", auto_adjust=False)
        exp1 = df['Close'].ewm(span=12, adjust=False).mean()
        exp2 = df['Close'].ewm(span=26, adjust=False).mean()
        macd = exp1 - exp2
        signal = macd.ewm(span=9, adjust=False).mean()
        return round(float(signal.values[-1]), 2)  # Extract raw float and round
    except:
        return None

def get_200d_50d_crossover(ticker):
    try:
        df = yf.download(ticker, period="1y", interval="1d", auto_adjust=False)
        print(f"{ticker}: Downloaded {len(df)} rows") 
        if len(df) < 200:
            print(f"{ticker}: Not enough data for 200D/50D crossover")
            return None
        ma50 = df['Close'].rolling(window=50).mean()
        print(f"{ticker} Last 10 MA50:\n{ma50.tail(10)}")
        ma200 = df['Close'].rolling(window=200).mean()
        print(f"{ticker} Last 10 MA200:\n{ma200.tail(10)}")
        for i in range(-5, 0):
            try:
                prev_50 = ma50.iloc[i-1].item()
                prev_200 = ma200.iloc[i-1].item()
                curr_50 = ma50.iloc[i].item()
                curr_200 = ma200.iloc[i].item()
            except:
                continue
            if pd.notna(prev_50) and pd.notna(prev_200) and pd.notna(curr_50) and pd.notna(curr_200):
                if prev_50 < prev_200 and curr_50 >= curr_200:
                    return 'Golden Cross'
                elif prev_50 > prev_200 and curr_50 <= curr_200:
                    return 'Death Cross'
        return 'No Cross'
    except Exception as e:
        return None

# Main logic (unchanged)
tickers = ['AAPL', 'MSFT', 'GOOGL']
stock_data = []

for ticker in tickers:
    try:
        stock = yf.Ticker(ticker)
        info = stock.info

        today_price = info.get('currentPrice')
        price_1mo = get_price_n_days_ago(stock, 30)
        price_3mo = get_price_n_days_ago(stock, 90)
        price_1y = get_price_n_days_ago(stock, 365)

        if today_price and price_1y:
            y_change = ((today_price - price_1y) / price_1y) * 100
        else:
            y_change = None

        data = {
            'Ticker': ticker,
            "Today's Share Price": today_price,
            'Price 1 Month Ago': price_1mo,
            'Price 3 Month Ago': price_3mo,
            '1Y % Price Change': round(y_change, 2) if y_change else None,
            '52-Week High': info.get('fiftyTwoWeekHigh'),
            '52-Week Low': info.get('fiftyTwoWeekLow'),
            'P/E Ratio': info.get('trailingPE'),
            'EV/EBITDA': info.get('enterpriseToEbitda'),
            'EV/Sales': info.get('enterpriseToRevenue'),
            'EBITDA': info.get('ebitda'),
            'Operating Margin': info.get('operatingMargins'),
            'Return on Equity (ROE)': info.get('returnOnEquity'),
            'Dividend Yield': info.get('dividendYield'),
            'Beta': info.get('beta'),
            'Market Cap': info.get('marketCap'),
            'RSI (14-day)': get_rsi(ticker),          # Now returns float
            'MACD Signal': get_macd_signal(ticker),    # Now returns float
            '200D/50D Crossover': get_200d_50d_crossover(ticker),
            'EPS Estimate (Next Year)': info.get('forwardEps'),
            '5-Year EPS Growth Forecast': info.get('earningsQuarterlyGrowth'),
            'Analyst Target Price': info.get('targetMeanPrice'),
            'ESG Score (Sustainalytics)': info.get('esgScores', {}).get('totalEsg') if info.get('esgScores') else None
        }

        if any(v is not None for v in data.values()):
            stock_data.append(data)

    except Exception as e:
        print(f"[{ticker}] Error: {e}")

# Save to Excel
df = pd.DataFrame(stock_data)
file_path = r"D:\omnidatax\company\Company Details\Upwork Project\Dr.Arthur Demo\Dr.Arthur_Demo\arthur_stock_data_detailed.xlsx"
df.to_excel(file_path, index=False)