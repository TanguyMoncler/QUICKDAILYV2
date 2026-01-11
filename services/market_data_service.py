import yfinance as yf
import pandas as pd

from data.universe import CAC_SMALL_90, CAC_MID_60, MS190, CUSTOM_LIST


def get_full_universe():
    universe = set(CAC_SMALL_90 + CAC_MID_60 + MS190 + CUSTOM_LIST)
    return list(universe)


def get_market_data():
    """
    Récupère les données de marché pour le tableau 1
    Retourne :
    - most_active (top 5 par volume)
    - best_performers (top 5 par % change)
    - worst_performers (bottom 5 par % change)
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

            results.append({
                "ticker": ticker,
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
