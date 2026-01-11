import requests
from bs4 import BeautifulSoup


def get_cac_small_90():
    """
    Récupère les composants du CAC Small 90 via l'API Euronext
    Retourne une liste de tickers yfinance (.PA)
    """
    url = "https://live.euronext.com/api/v1/indices/FR0003999481-XPAR/components"

    response = requests.get(url, timeout=10)
    response.raise_for_status()

    data = response.json()
    tickers = []

    for item in data.get("components", []):
        symbol = item.get("symbol", "").strip()
        if symbol:  # éviter symboles vides
            tickers.append(f"{symbol}.PA")

    return tickers


def get_cac_mid_90():
    """
    Scrape les composants du CAC Mid 90
    Retourne une liste de tickers yfinance
    """
    pass


def get_ms190():
    """
    Scrape les composants du MS190
    Retourne une liste de tickers yfinance
    """
    pass


def get_full_universe(custom_list=None):
    """
    Combine :
    - CAC Small 90
    - CAC Mid 90
    - MS190
    - custom_list (optionnelle)

    Retourne une liste UNIQUE de tickers yfinance
    """
    if custom_list is None:
        custom_list = []

    universe = set()

    universe.update(get_cac_small_90())
    universe.update(get_cac_mid_90())
    universe.update(get_ms190())
    universe.update(custom_list)

    return list(universe)