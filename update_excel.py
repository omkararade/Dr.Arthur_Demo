import sys                                                      # Built-in module: lets us access and modify Python runtime settings (like sys.path)
import os                                                       # Built-in module: work with file paths, directories, environment
import xlwings as xw                                            # xlwings: control Excel from Python (read/write/format workbooks)
import pandas as pd                                             # pandas: DataFrame handling for tabular data



# ------✅ Display settings for clean DataFrame in terminal------
pd.set_option('display.max_columns', None)                      # Show all columns when printing DataFrames (no truncation)
pd.set_option('display.width', None)                            # Let pandas use full terminal width (prevents wrapping)


# ------✅ Add project root and scripts directory to Python path------
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))       #Get absolute path to the project root: start at this file's folder, then go one level up
scripts_dir = os.path.join(project_root, "scripts")                                 #Build path to the "scripts" subfolder inside the project root

if project_root not in sys.path:                                                    # If project root isn't already in Python's import search paths...
    sys.path.append(project_root)                                                   # ...add it so imports can find modules there
if scripts_dir not in sys.path:                                                     # If "scripts" folder isn't already in Python's import paths...
    sys.path.append(scripts_dir)                                                    # ...add it too

try:
    from fetch_data import fetch_stock_data_with_indicators, get_tickers_from_excel #Try importing the two helper functions from fetch_data.py (expected to be in project root or scripts)
    print("✅ Successfully imported fetch_data functions")
except ImportError:                                                                 #If the import fails (wrong path/name), handle gracefully
    print("❌ Could not import fetch_data. Check file location/import path.")
    sys.exit()                                                                      # Exit the program since we can't continue without these functions

def format_excel(sheet):
    """ ✅ Apply Excel Formatting: header, autofit, borders, freeze pane, alignment """
    print("🎨 Applying Excel formatting...")
    
    # Header Range
    header = sheet.range("A1").expand("right")                                      # Select header row starting at A1 and expanding to the last used column (first row only)
    header.api.Font.Bold = True                                                     # Make header text bold via Excel COM API
    header.api.Interior.Color = 0xD9E1F2                                            # Fill header background with a light blue color (BGR integer)

    # All Data Range (Body + Header)
    last_row = sheet.range("A1").expand("down").last_cell.row                       # Find the last used row by expanding downward from A1 and reading the last cell's row index
    last_col = sheet.range("A1").expand("right").last_cell.column                   # Find the last used column by expanding rightward from A1 and reading the last cell's column index
    used_range = sheet.range((1, 1), (last_row, last_col))                          # Create a range covering the full used area from A1 to bottom-right cell
    

    # Borders Around All Cells
    for border_id in range(7, 13):                                                  # Excel's border ids 7..12 map to edges/inside borders (xlEdgeLeft=7 ... xlInsideHorizontal=12)
        used_range.api.Borders(border_id).LineStyle = 1                             # Set each border line style to continuous (xlContinuous=1)

    # Center Align
    used_range.api.HorizontalAlignment = -4108                                      # Center text horizontally (xlCenter = -4108)
    

    # Auto-fit Columns
    sheet.range("A1").expand().columns.autofit()                                    # Auto-fit all columns in the contiguous used block starting at A1
    

    # Freeze Top Header Row
    sheet.api.Application.ActiveWindow.SplitRow = 1                                 # Split the window after the first row (so row 1 can be frozen)
    sheet.api.Application.ActiveWindow.FreezePanes = True                           # Freeze panes so header row stays visible during scrolling

    print("✅ Excel formatting applied successfully!")



def update_excel():
    try:
        print("🔄 Starting Excel update process...")
        base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))   # Resolve the absolute path to the project base directory (one level above this script)
        excel_path = os.path.join(base_dir, "dashboard.xlsm")                       # Build full path to the target Excel workbook (macro-enabled)
        print(f"📁 Excel file path: {excel_path}")

        if not os.path.exists(excel_path):                                          # If the Excel file is missing, stop early
            print(f"❌ dashboard.xlsm not found!")
            return False

        app = xw.apps.active                                                        # Get the currently running Excel instance (if any)
        if app is None:                                                             # If Excel isn't open, we prefer the user to open it first (or we could open it ourselves)
            print("❌ No active Excel instance detected. Please open the Excel file first.")
            return False

        # Find the opened Excel workbook
        workbook = None
        for wb in app.books:                                                        # Iterate over all open workbooks in the active Excel app
            if "dashboard.xlsm" in wb.name:                                         # Match by filename; if found, keep a reference
                workbook = wb

        if workbook is None:                                                        # If workbook isn't already open, open it now
            print("⚠️ Workbook is not open. Opening it now...")
            workbook = app.books.open(excel_path)
        else:
            print(f"✅ Found open workbook: {workbook.name}")

        # ✅ Select the sheet where data will be written
        sheet = workbook.sheets["RawData"]                                          # Get a handle to the "RawData" sheet (must exist in the workbook)

        # ✅ Read tickers from Sheet1
        print("📊 Fetching tickers from Sheet1...")
        tickers = get_tickers_from_excel(sheet_name="Sheet1")                       # Use helper to read tickers from the first column of the "Sheet1" sheet
        print("✅ Tickers found:", tickers)

        # ✅ Fetch stock data
        print("📈 Fetching stock data...")
        df = fetch_stock_data_with_indicators(tickers)                              # Download yfinance data + compute indicators (RSI, MACD) → return a DataFrame

        # print("\n📄 Data Preview:")                                               # Pretty-printed DataFrame in terminal (using earlier pd.set_option settings)
        # print(df)  

        # ✅ Clear and write data to RawData sheet
        print("💾 Writing data to RawData sheet...")
        sheet.clear()                                                               # Clear the entire sheet to remove old content/formatting
        sheet.range("A1").value = df                                                # Write the DataFrame starting at cell A1 (headers + data)

        # ✅ Apply formatting
        format_excel(sheet)                                                         # Call our formatter to style the sheet (headers, borders, freeze, autofit)

        print("✅ Excel updated successfully!")                                    # Indicate success to the caller
        return True
        

    except Exception as e:                                                          # Catch any runtime errors and print a helpful trace
        print(f"❌ Error in update_excel: {e}")
        import traceback
        traceback.print_exc()
        return False                                                                # Indicate failure to the caller
        

if __name__ == "__main__":
    update_excel()                                                                  # If this script is run directly (not imported), execute the update workflow