import yfinance as yf
import pandas as pd

from data.universe import CUSTOM_LIST


def get_full_universe():
    # Pour l'instant, on utilise la liste CUSTOM_LIST. Vous pouvez étendre
    # pour inclure CAC_SMALL_90, CAC_MID_60, MS190 si ces listes sont ajoutées.
    return list(set(CUSTOM_LIST))


def get_market_data():
    """
    Récupère les données de marché pour le tableau 1
    Retourne :
    - most_active (top 5 par volume)
    - best_performers (top 5 par % change)
    - worst_performers (bottom 5 par % change)
    Chaque enregistrement contient maintenant également le nom court (shortName) si disponible.
    """

    tickers = get_full_universe()

    data = yf.download(
        tickers=tickers,
        period="1d",
        interval="1d",
        group_by="ticker",
        auto_adjust=False,
        threads=True,
        progress=False
    )

    results = []

    for ticker in tickers:
        try:
            df = data[ticker]
            open_price = df["Open"].iloc[0]
            close_price = df["Close"].iloc[0]
            volume = df["Volume"].iloc[0]

            if open_price == 0 or pd.isna(open_price):
                continue

            pct_change = ((close_price - open_price) / open_price) * 100

            # Récupération du nom via yfinance si possible
            name = None
            try:
                info = yf.Ticker(ticker).info
                name = info.get("shortName") or info.get("longName")
            except Exception:
                name = None

            results.append({
                "ticker": ticker,
                "name": name or ticker.replace('.PA', ''),
                "open": open_price,
                "close": close_price,
                "volume": volume,
                "pct_change": pct_change
            })

        except Exception:
            continue

    df_results = pd.DataFrame(results)

    most_active = (
        df_results
        .sort_values("volume", ascending=False)
        .head(5)
        .to_dict("records")
    )

    best_performers = (
        df_results
        .sort_values("pct_change", ascending=False)
        .head(5)
        .to_dict("records")
    )

    worst_performers = (
        df_results
        .sort_values("pct_change", ascending=True)
        .head(5)
        .to_dict("records")
    )

    return {
        "most_active": most_active,
        "best_performers": best_performers,
        "worst_performers": worst_performers
    }


def _fetch_single_ticker(ticker):
    """Retourne (close, pct_change) pour un ticker"""
    try:
        df = yf.download(tickers=[ticker], period="1d", interval="1d", progress=False)
        if ticker in df and not df.empty:
            row = df[ticker].iloc[0]
            open_price = row["Open"]
            close_price = row["Close"]
        else:
            # Certains tickers retournent un DataFrame sans niveau de ticker
            if "Open" in df.columns:
                open_price = df["Open"].iloc[0]
                close_price = df["Close"].iloc[0]
            else:
                return None

        if open_price == 0 or pd.isna(open_price):
            return None

        pct_change = ((close_price - open_price) / open_price) * 100
        return {
            "close": close_price,
            "pct_change": pct_change
        }
    except Exception:
        return None


def get_table_2_data(table_2_tickers):
    """Retourne une liste d'éléments pour le tableau 2: placeholder, display, pct_change"""
    results = []
    for placeholder, ticker in table_2_tickers:
        info = _fetch_single_ticker(ticker)
        if info:
            display = f"{info['close']:.2f}"
            results.append({
                "placeholder": placeholder,
                "display": display,
                "pct_change": info["pct_change"]
            })
        else:
            results.append({
                "placeholder": placeholder,
                "display": "-",
                "pct_change": 0.0
            })

    return results


def get_table_3_data(sector_mapping):
    """Retourne une liste d'éléments pour le tableau 3: placeholder, display, pct_change
    sector_mapping est un dict {placeholder: ticker_or_None}
    """
    results = []
    for placeholder, ticker in sector_mapping.items():
        if not ticker:
            results.append({
                "placeholder": placeholder,
                "display": "-",
                "pct_change": 0.0
            })
            continue

        info = _fetch_single_ticker(ticker)
        if info:
            display = f"{info['close']:.2f}"
            results.append({
                "placeholder": placeholder,
                "display": display,
                "pct_change": info["pct_change"]
            })
        else:
            results.append({
                "placeholder": placeholder,
                "display": "-",
                "pct_change": 0.0
            })

    return results
