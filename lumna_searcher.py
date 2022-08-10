import pandas as pd


def get_price(file_name="C:\\Users\\User1\\Desktop\\OrderMakerDocuments\\ProviderPrices\\LumnaPrice.xls"):
    price = pd.read_excel(file_name, header=3, usecols=list(range(2, 13)))
    return price


def search(price: pd.DataFrame, formalized_order):
    for order_row in formalized_order:
        isbn = "".join(order_row["ISBN"].split("-"))

        for price_row in price.iterrows():
            price_data = price_row[1]
            if not pd.isnull(price_data["EAN"]) and int(price_data["EAN"]) == int(isbn):
                order_row["Люмн. Код"] = price_data["Код"]
                order_row["Люмн. Остаток"] = price_data["Остаток"]
                order_row["Люмн. Цена"] = price_data["Цена"]
                order_row["Люмн. Цена со скидкой"] = price_data["Цена со скидкой"]
                break

    return formalized_order


