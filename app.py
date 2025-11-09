import streamlit as st
import pandas as pd
import yfinance as yf
import os

TICKER_FILE = "tickers.txt"  # file to store ticker symbols

# Read tickers from file
def load_tickers():
    if os.path.exists(TICKER_FILE):
        with open(TICKER_FILE, "r") as file:
            tickers = file.read().strip()
            return [t.strip().upper() for t in tickers.split(",") if t.strip()]
    return ["AAPL", "MSFT", "TSLA", "GOOGL"]  # default if file missing

# Save tickers to file
def save_tickers(tickers):
    with open(TICKER_FILE, "w") as file:
        file.write(",".join(tickers))

# RSI Calculation
def calculate_rsi(close, period=14):
    delta = close.diff()
    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)
    avg_gain = gain.rolling(window=period).mean()
    avg_loss = loss.rolling(window=period).mean()
    rs = avg_gain / avg_loss
    return 100 - (100 / (1 + rs))

# MACD
def calculate_macd(close, short=12, long=26, signal=9):
    ema_short = close.ewm(span=short, adjust=False).mean()
    ema_long = close.ewm(span=long, adjust=False).mean()
    macd = ema_short - ema_long
    signal_line = macd.ewm(span=signal, adjust=False).mean()
    return macd, signal_line

st.title("📊 Persistent Stock Dashboard (RSI, MACD, Price Data)")

# Load saved tickers
tickers = load_tickers()

# --- Ticker Input Box for adding new stock ---
new_ticker = st.text_input("➕ Add Ticker Symbol (Example: NFLX, NVDA):")

if st.button("Add Ticker"):
    if new_ticker.strip().upper() not in tickers:
        tickers.append(new_ticker.strip().upper())
        save_tickers(tickers)
        st.success(f"Added {new_ticker.upper()}")
        st.rerun()
    else:
        st.warning("Ticker already exists.")

# --- Automatically fetch stock data when app loads ---
all_data = []
for ticker in tickers:
    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        hist = stock.history(period="6mo")

        rsi = macd_val = signal = None
        if not hist.empty:
            close = hist["Close"]
            rsi = calculate_rsi(close).iloc[-1] if len(close) > 14 else None
            macd_series, signal_series = calculate_macd(close)
            macd_val = macd_series.iloc[-1]
            signal = signal_series.iloc[-1]

        all_data.append({
            "Ticker": ticker,
            "Price": round(info.get("regularMarketPrice", 0), 2),
            "P/E Ratio": info.get("trailingPE", "N/A"),
            "Market Cap (B)": round(info.get("marketCap", 0) / 1e9, 2) if info.get("marketCap") else "N/A",
            "Dividend Yield": info.get("dividendYield", "N/A"),
            "RSI (14)": round(rsi, 2) if rsi else "N/A",
            "MACD": round(macd_val, 2) if macd_val else "N/A",
            "Signal": round(signal, 2) if signal else "N/A",
        })
    except:
        all_data.append({"Ticker": ticker, "Error": "Failed to fetch"})

# Show Table
df = pd.DataFrame(all_data)
st.dataframe(df, use_container_width=True)

# Download Button
st.download_button("⬇ Download CSV", df.to_csv(index=False), "stock_data.csv")
