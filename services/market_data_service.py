import yfinance as yf
import pandas as pd
import warnings
from datetime import datetime

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

    def _get_daily_data_single_row(self, ticker):
        """Get daily data, accepting single row (for indices like EURO STOXX)"""
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            # For indices, fetch longer history to get at least 2 trading days
            data = yf.download(ticker, period="60d", interval="1d", progress=False)
        
        if data is None or len(data) < 1:
            self.unavailable_tickers.append(ticker)
            return None
        
        # If we got 2+ rows, return previous and current
        if len(data) >= 2:
            return data.iloc[-2], data.iloc[-1]
        
        # If only 1 row (common for indices), return it with itself as fallback
        # The calling code will detect this and use Open vs Close instead
        return data.iloc[-1], data.iloc[-1]

    def _get_intraday_and_avg_volume(self, ticker):
        """Get today's intraday volume sum and 10-day average volume"""
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            # Get 1-minute intraday data for today
            intraday = yf.download(ticker, period="1d", interval="1m", progress=False)
            
            # Get last 10 days of daily data for average volume
            daily = yf.download(ticker, period="11d", interval="1d", progress=False)
        
        if intraday is None or len(intraday) == 0:
            return None, None
        
        if daily is None or len(daily) < 10:
            return None, None
        
        try:
            # Get volume column (handle both single and multi-index columns)
            intraday_volume = intraday["Volume"] if "Volume" in intraday.columns else intraday[("Volume", ticker)]
            daily_volume = daily["Volume"] if "Volume" in daily.columns else daily[("Volume", ticker)]
            
            # Sum today's intraday volume
            today_volume = intraday_volume.sum().item()
            
            # Calculate average volume for the last 10 trading days (excluding today)
            avg_volume_10d = daily_volume.iloc[:-1].tail(10).mean().item()
            
            return today_volume, avg_volume_10d
        except (KeyError, TypeError):
            return None, None

    def compute_table_1(self, universe):
        rows = []

        for item in universe:
            data = self._get_daily_data(item["Ticker"])
            if data is None:
                continue
            prev, curr = data

            # Convert Series to scalars (handle both single and multi-index columns)
            try:
                close_curr = float(curr["Close"].item() if hasattr(curr["Close"], 'item') else curr["Close"])
                close_prev = float(prev["Close"].item() if hasattr(prev["Close"], 'item') else prev["Close"])
            except (KeyError, TypeError):
                try:
                    close_curr = float(curr[("Close", item["Ticker"])].item())
                    close_prev = float(prev[("Close", item["Ticker"])].item())
                except (KeyError, TypeError):
                    continue
            
            variation = (close_curr - close_prev) / close_prev
            multiple = close_curr / close_prev
            
            # Get intraday volume and 10-day average
            today_volume, avg_volume_10d = self._get_intraday_and_avg_volume(item["Ticker"])
            
            if today_volume is None or avg_volume_10d is None:
                volume_multiple = 1.0
            else:
                volume_multiple = today_volume / avg_volume_10d if avg_volume_10d > 0 else 1.0

            rows.append({
                "name": item["Name"],
                "variation": variation,
                "multiple": multiple,
                "volume_multiple": volume_multiple
            })

        if not rows:
            return {
                "most_active": pd.DataFrame(),
                "best": pd.DataFrame(),
                "worst": pd.DataFrame()
            }

        df = pd.DataFrame(rows)
        df = df.reset_index(drop=True)

        most_active = df.nlargest(5, "volume_multiple")
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
        results = []

        for item in universe:
            ticker = item["Ticker"]
            data = self._get_daily_data_single_row(ticker)
            if data is None:
                continue
            prev, curr = data

            try:
                close = curr["Close"].item() if hasattr(curr["Close"], 'item') else float(curr["Close"])
                prev_close = prev["Close"].item() if hasattr(prev["Close"], 'item') else float(prev["Close"])
            except (KeyError, TypeError):
                try:
                    close = float(curr[("Close", ticker)].item())
                    prev_close = float(prev[("Close", ticker)].item())
                except (KeyError, TypeError):
                    continue
            
            if close == prev_close:
                # If same close, calculate intra-day variation (Open vs Close)
                try:
                    open_price = curr["Open"].item() if hasattr(curr["Open"], 'item') else float(curr["Open"])
                except (KeyError, TypeError):
                    try:
                        open_price = float(curr[("Open", ticker)].item())
                    except (KeyError, TypeError):
                        variation = None
                    else:
                        if open_price != 0 and open_price != close:
                            variation = (close - open_price) / open_price
                        else:
                            variation = None
                else:
                    if open_price != 0 and open_price != close:
                        variation = (close - open_price) / open_price
                    else:
                        variation = None
            else:
                if prev_close != 0:
                    variation = (close - prev_close) / prev_close
                else:
                    variation = None

            results.append({
                "ticker": ticker,
                "close": close,
                "variation": variation
            })

        return results

    def _compute_simple_table(self, universe):
        results = []

        for item in universe:
            ticker = item["Ticker"]
            data = self._get_daily_data(ticker)
            if data is None:
                continue
            prev, curr = data

            try:
                close = curr["Close"].item() if hasattr(curr["Close"], 'item') else float(curr["Close"])
                prev_close = prev["Close"].item() if hasattr(prev["Close"], 'item') else float(prev["Close"])
            except (KeyError, TypeError):
                try:
                    close = float(curr[("Close", ticker)].item())
                    prev_close = float(prev[("Close", ticker)].item())
                except (KeyError, TypeError):
                    continue
            
            if close == prev_close:
                variation = None
            else:
                if prev_close != 0:
                    variation = (close - prev_close) / prev_close
                else:
                    variation = None

            results.append({
                "ticker": ticker,
                "close": close,
                "variation": variation
            })

        return results