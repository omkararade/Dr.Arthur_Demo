import pandas as pd

def calculate_rsi(close_prices, period=14):
    delta = close_prices.diff()

    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)

    avg_gain = gain.rolling(window=period).mean()
    avg_loss = loss.rolling(window=period).mean()

    rs = avg_gain / avg_loss
    rsi = 100 - (100 / (1 + rs))

    return rsi



def calculate_macd(close_prices, short_window=12, long_window=26, signal_window=9):
    ema_short = close_prices.ewm(span=short_window, adjust=False).mean()
    ema_long = close_prices.ewm(span=long_window, adjust=False).mean()

    macd = ema_short - ema_long
    signal = macd.ewm(span=signal_window, adjust=False).mean()
    histogram = macd - signal

    return macd, signal, histogram
