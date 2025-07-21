import streamlit as st
import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
import subprocess
import os

# Helper functions (same as test.py)
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
        return float(rsi.values[-1])
    except:
        return None

def get_macd_signal(ticker):
    try:
        df = yf.download(ticker, period="6mo", interval="1d", auto_adjust=False)
        exp1 = df['Close'].ewm(span=12, adjust=False).mean()
        exp2 = df['Close'].ewm(span=26, adjust=False).mean()
        macd = exp1 - exp2
        signal = macd.ewm(span=9, adjust=False).mean()
        return round(float(signal.values[-1]), 2)
    except:
        return None

def get_200d_50d_crossover(ticker):
    try:
        df = yf.download(ticker, period="1y", interval="1d", auto_adjust=False)
        if len(df) < 200:
            return None
        ma50 = df['Close'].rolling(window=50).mean()
        ma200 = df['Close'].rolling(window=200).mean()
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
    
with st.sidebar:
    st.header("📋 Instructions")
    
    st.subheader("File Upload Requirements")
    st.markdown("""
    1. **File Format**: Excel (.xlsx or .xls)
    2. **Required Column**: 
       - Must have a column named **'Ticker'** (case-sensitive)
    3. **Example Structure**:
    """)
    
    # Example dataframe
    example_df = pd.DataFrame({
        'Ticker': ['AAPL', 'MSFT', 'GOOGL'],
        'Company Name': ['Apple Inc.', 'Microsoft', 'Alphabet Inc.'],
        'Notes': ['Tech', 'Software', 'Parent of Google']
    })
    st.dataframe(example_df)
    
    st.markdown("""
    4. **How to Upload**:
       - Click "Browse files" or drag-and-drop your file
       - Ensure only one sheet contains your ticker data
       - Wait for the confirmation message
    """)
    
    st.subheader("Analysis Options")
    st.markdown("""
      
    - Quick In-App Analysis
      - Faster results directly in the app
      - Limited to key metrics
      - Best for quick checks
    """)
    
    st.subheader("Need Help?")
    st.markdown("""
    - Ensure your Excel file isn't open while uploading
    - Tickers must be valid stock symbols
    - For large files (>50 tickers), use Option 1
    - Contact omkararade@gmail.com for assistance
    """)

# Streamlit UI
st.title("📈 Stock Analysis Tool")

# Option 1: Run the test.py script and show output
# st.header("Option 1: Run Full Analysis (test.py)")
# if st.button("Run Full Analysis"):
#     # Execute the test.py script
#     result = subprocess.run(["python", "test.py"], capture_output=True, text=True)
    
#     # Display output
#     st.subheader("Analysis Output")
#     st.text(result.stdout)
    
#     # Show errors if any
#     if result.stderr:
#         st.error("Errors encountered:")
#         st.text(result.stderr)
    
#     # Verify and display the Excel file
#     file_path = r"D:\omnidatax\company\Company Details\Upwork Project\Dr.Arthur Demo\Dr.Arthur_Demo\arthur_stock_data_detailed.xlsx"
#     if os.path.exists(file_path):
#         st.success("Analysis completed successfully!")
#         df = pd.read_excel(file_path)
#         st.dataframe(df)
        
#         # Add download button
#         with open(file_path, "rb") as file:
#             st.download_button(
#                 label="Download Excel Report",
#                 data=file,
#                 file_name="arthur_stock_data_detailed.xlsx",
#                 mime="application/vnd.ms-excel"
#             )
#     else:
#         st.error(f"Output file not found at: {file_path}")

# Option 2: In-app analysis
st.header("Option 2: Analyze in App")
uploaded_file = st.file_uploader("Upload Excel file with tickers", type=["xlsx", "xls"])

if uploaded_file:
    # Read Excel file
    df = pd.read_excel(uploaded_file)
    
    # Check if 'Ticker' column exists
    if 'Ticker' not in df.columns:
        st.error("Excel file must contain a 'Ticker' column")
    else:
        tickers = df['Ticker'].tolist()
        st.success(f"Found {len(tickers)} tickers: {', '.join(tickers)}")
        
        # Analyze button
        if st.button("Analyze Stocks"):
            # Create progress elements
            status_container = st.empty()
            progress_bar = st.progress(0)
            results = []
            
            # Start analysis
            status_container.subheader("🚀 Starting analysis...")
            
            for i, ticker in enumerate(tickers):
                try:
                    # Update progress
                    progress = int((i + 1) / len(tickers) * 100)
                    progress_bar.progress(progress)
                    status_container.subheader(f"🔍 Analyzing {ticker} ({i+1}/{len(tickers)})...")
                    
                    stock = yf.Ticker(ticker)
                    info = stock.info

                    today_price = info.get('currentPrice')
                    price_1mo = get_price_n_days_ago(stock, 30)
                    price_1y = get_price_n_days_ago(stock, 365)
                    y_change = ((today_price - price_1y) / price_1y) * 100 if (today_price and price_1y) else None

                    # Full data collection
                    data = {
                        'Ticker': ticker,
                        "Today's Share Price": today_price,
                        'Price 1 Month Ago': price_1mo,
                        'Price 3 Months Ago': get_price_n_days_ago(stock, 90),
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
                        'RSI (14-day)': get_rsi(ticker),
                        'MACD Signal': get_macd_signal(ticker),
                        '200D/50D Crossover': get_200d_50d_crossover(ticker),
                        'EPS Estimate (Next Year)': info.get('forwardEps'),
                        '5-Year EPS Growth Forecast': info.get('earningsQuarterlyGrowth'),
                        'Analyst Target Price': info.get('targetMeanPrice'),
                        'ESG Score': info.get('esgScores', {}).get('totalEsg') if info.get('esgScores') else None
                    }
                    results.append(data)
                    
                except Exception as e:
                    st.warning(f"Error analyzing {ticker}: {str(e)}")
            
            # Clear progress elements
            status_container.empty()
            progress_bar.empty()
            
            # Display results
            st.subheader("✅ Analysis Complete!")
            results_df = pd.DataFrame(results)
            st.dataframe(results_df)
            
            # Add download button
            if not results_df.empty:
                csv = results_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download CSV",
                    data=csv,
                    file_name="stock_analysis.csv",
                    mime="text/csv"
                )