from datetime import datetime
from services.market_data_service import MarketDataService
from services.word_service import WordService
from data.universe import LIST_1, LIST_2, LIST_3

TEMPLATE_PATH = "templates/closing_template.docx"
OUTPUT_DIR = "output"

def main():
    today = datetime.today()
    file_date = today.strftime("%Y-%m-%d")
    display_date = today.strftime("%B %d, %Y").replace(" 0", " ")
    display_date = display_date.replace("1,", "1st,").replace("2,", "2nd,").replace("3,", "3rd,")
    display_date = display_date.replace("21,", "21st,").replace("22,", "22nd,").replace("23,", "23rd,")
    display_date = display_date.replace("31,", "31st,")

    output_file = f"{OUTPUT_DIR}/{file_date} daily closing.docx"

    market_service = MarketDataService()
    word_service = WordService(TEMPLATE_PATH)

    table1_data = market_service.compute_table_1(LIST_1)
    table2_data = market_service.compute_table_2(LIST_2)
    table3_data = market_service.compute_table_3(LIST_3)

    word_service.replace_date(display_date)
    word_service.fill_table_1(table1_data)
    word_service.fill_table_2(table2_data)
    word_service.fill_table_3(table3_data)

    word_service.save(output_file)

    if market_service.unavailable_tickers:
        print("\n✗ Unavailable tickers (skipped):")
        for ticker in market_service.unavailable_tickers:
            print(f"  - {ticker}")
    else:
        print("\n✓ All tickers loaded successfully")
    
    print("\nDONE !")

if __name__ == "__main__":
    main()