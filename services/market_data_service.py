import yfinance as yf
import pandas as pd
import warnings

class MarketDataService:

    def __init__(self):
        self.unavailable_tickers = []

    def _get_daily_data(self, ticker):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            data = yf.download(ticker, period="2d", interval="1d", progress=False)
        
        if data is None or len(data) < 2:
            self.unavailable_tickers.append(ticker)
            return None
        return data.iloc[-2], data.iloc[-1]

    def compute_table_1(self, universe):
        rows = []

        for item in universe:
            data = self._get_daily_data(item["Ticker"])
            if data is None:
                continue
            prev, curr = data

            # Convert Series to scalars
            close_curr = float(curr["Close"].item() if hasattr(curr["Close"], 'item') else curr["Close"])
            close_prev = float(prev["Close"].item() if hasattr(prev["Close"], 'item') else prev["Close"])
            volume = float(curr["Volume"].item() if hasattr(curr["Volume"], 'item') else curr["Volume"])
            
            variation = (close_curr - close_prev) / close_prev
            multiple = close_curr / close_prev

            rows.append({
                "name": item["Name"],
                "variation": variation,
                "multiple": multiple,
                "volume": volume
            })

        if not rows:
            return {
                "most_active": pd.DataFrame(),
                "best": pd.DataFrame(),
                "worst": pd.DataFrame()
            }

        df = pd.DataFrame(rows)
        df = df.reset_index(drop=True)

        most_active = df.nlargest(5, "volume")
        best = df.nlargest(5, "variation")
        worst = df.nsmallest(5, "variation")

        return {
            "most_active": most_active,
            "best": best,
            "worst": worst
        }

    def compute_table_2(self, universe):
        return self._compute_simple_table(universe)

    def compute_table_3(self, universe):
        return self._compute_simple_table(universe)

    def _compute_simple_table(self, universe):
        results = []

        for item in universe:
            ticker = item["Ticker"]
            data = self._get_daily_data(ticker)
            if data is None:
                continue
            prev, curr = data

            close = curr["Close"].item() if hasattr(curr["Close"], 'item') else float(curr["Close"])
            prev_close = prev["Close"].item() if hasattr(prev["Close"], 'item') else float(prev["Close"])
            
            if close == prev_close:
                variation = None
            else:
                variation = (close - prev_close) / prev_close

            results.append({
                "ticker": ticker,
                "close": close,
                "variation": variation
            })

        return results