from services.market_data_service import get_market_data

if __name__ == "__main__":
    data = get_market_data()

    print("MOST ACTIVE")
    for x in data["most_active"]:
        print(x)

    print("\nBEST PERFORMERS")
    for x in data["best_performers"]:
        print(x)

    print("\nWORST PERFORMERS")
    for x in data["worst_performers"]:
        print(x)
