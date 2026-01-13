QUICKDAILYV2

This project generates a daily closing Word report using data from Yahoo Finance (via yfinance).

Usage
- Install dependencies: pip install -r requirements.txt
- Configure tickers (optional): edit `config/settings.py` to set `TABLE_2_TICKERS` and `SECTOR_TICKERS`.
- Run: python main.py
- Output: `output/YYYY-MM-DD daily closing.docx` (date is the day when the script is run).

Notes
- The template is `templates/closing_template.docx` and must contain placeholders such as `{{DATE}}`, table placeholders described in the repository.
- For Table 3, set `SECTOR_TICKERS` to appropriate yfinance tickers for each sector; if left as None the report will show `-` for those cells.
