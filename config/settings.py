# Mapping for Table 2: (placeholder, yfinance_ticker)
TABLE_2_TICKERS = [
    ("{{^STOXX50E}}", "^STOXX50E"),
    ("{{^GDAXI}}", "^GDAXI"),
    ("{{^FCHI}}", "^FCHI"),
    ("{{FTSEMIB.MI}}", "FTSEMIB.MI"),
    ("{{^VIX}}", "^VIX"),
    ("{{^EURUSD=X}}", "EURUSD=X"),
    ("{{GC=F}}", "GC=F"),
    ("{{CL=F}}", "CL=F"),
    ("{{BTC-EUR}}", "BTC-EUR"),
]

# Mapping for Table 3: (placeholder, yfinance_ticker)
# Replace these with the actual tickers you want to track for each sector.
SECTOR_TICKERS = {
    # STOXX sector tickers (Yahoo-style symbols)
    "{{Auto}}": "^SXAP",            # Automobiles & Parts
    "{{Banks}}": "^SX7E",           # Banks
    "{{Basic Resources}}": "^SXPP", # Basic Resources
    "{{Chemicals}}": "^SXCH",       # Chemicals
    "{{Food & Beverages}}": "^SX3P",# Food & Beverages
}

# Display / formatting settings
NAME_COLOR_HEX = "0070C0"  # names and multiples color
GREEN_HEX = "00B050"
RED_HEX = "C00000"
BLACK_HEX = "000000"
FONT_NAME = "Sitka Display"
FONT_SIZE_PT = 8
