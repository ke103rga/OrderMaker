import pandas as pd
import re


def get_price(file_name="C:\\Users\\User1\\Desktop\\OrderMakerDocuments\\ProviderPrices\\interservicePrice.xlsx"):
    price = pd.read_excel(file_name, header=2)
    return price


def search(price: pd.DataFrame, order):

    def cls_and_part(cls, part, name):
        if cls and part:
            return f'{part}' in name \
                   and f'{cls}кл' in name
        elif cls:
            return f'{cls}кл' in name
        elif part:
            return f'{part}' in name
        else:
            return True

    def contains_key_words(name, author, publish, order_row):
        key_words = order_row.get("key_words")
        info = order_row.get("info")
        cls = order_row.get("cls")
        part = order_row.get("part")
        if all([word in name for word in key_words]) \
                and order_row.get("name")[:-2] in name \
                and all(word in name for word in info) \
                and order_row.get("author")[:-2] in name:
            if author and publish:
                if (order_row.get("author") in author or order_row.get("coauthor") in author) \
                        and order_row.get("publish") in publish:
                    return cls_and_part(cls, part, name)
            elif author:
                if order_row.get("author") in author or order_row.get("coauthor") in author:
                    return cls_and_part(cls, part, name)
            elif publish:
                if order_row.get("publish") in publish:
                    return cls_and_part(cls, part, name)
            else:
                return cls_and_part(cls, part, name)
        else:
            return False


    result = []

    for order_row in order:
        for row in price.iterrows():
            data = row[1]
            name = data["Наименование"].lower()
            if not pd.isnull(data["Автор"]):
                author = data["Автор"].lower()
            else:
                author = ""
            if not pd.isnull(data["Издательство"]):
                publish = data["Издательство"].lower()
            else:
                publish = ""

            if contains_key_words(name, author, publish, order_row):
                # if name.split()[0][:-2] in order_row.values():
                result.append(data.to_dict())
                print(order_row)

    # df = pd.DataFrame(result)
    # df.to_excel("TryResult.xls")
    return result


def parse_key_words(key_words):
    keys = ["Раб", "тет", "конт", "карт", "учеб", "атл", "проп"]
    cats = []

    def contains_digit(word):
        return re.search(r"\d", word)

    key_words = key_words.lower()
    # for word in key_words:
    #     if len(word) <= 3 or contains_digit(word):
    #         key_words.remove(word)
    # key_words = list(map(lambda word: word[:3].lower(), key_words))
    for key in keys:
        if key in key_words:
            cats.append(key)

    return cats


def formalize_search_result(result):
    columns_for_del = ["Страниц", "Заказ", "Сумма", "Сумма со скидкой"]

    for row in result:
        for column in columns_for_del:
            del row[column]

        row["Интер. Цена"] = row.pop("Цена")
        row["Интер. Цена со скидкой"] = row.pop("Цена со скидкой")

        row["Люмн. Цена"] = 0
        row["Люмн. Цена со скидкой"] = 0

    # df = pd.DataFrame(result)
    # df.to_excel("TryResult.xls")

    return result
