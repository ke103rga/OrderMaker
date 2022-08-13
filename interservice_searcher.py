import pandas as pd
import re
from buyer_order_converter import get_order, convert_order


def get_price(file_name="C:\\Users\\User1\\Downloads\\knigi-polnyy-prays_3.xlsx"):
    price = pd.read_excel(file_name, header=12)
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
            if not pd.isnull(data["Наименование"]):
                name = data["Наименование"].lower()
            else:
                continue
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
                data = data.to_dict()
                data["Количество"] = order_row.get("amount")

                result.append(data)
                print(order_row)

    # df = pd.DataFrame(result)
    # df.to_excel("TryResult.xls")
    return delete_clones(result)


def delete_clones(lst):
    used_dicts = []

    for dct in lst:
        if dct in used_dicts:
            continue
        else:
            used_dicts.append(dct)

    return used_dicts


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
    columns_for_del = ["Страниц", "Заказ", "Сумма", "Сумма со скидкой", "Автор", "Артикул", "Переплет", "Новинка", "В упаковке",]

    for row in result:
        for column in columns_for_del:
            del row[column]

        row["Заказ"] = ""
        row["Заказ (количество)"] = ""
        row["Проверка"] = ""

        row["Интер. Цена"] = row.pop("Цена")
        row["Интер. Цена со скидкой"] = row.pop("Цена со скидкой")
        row["Интер. Остаток"] = row.pop("Остаток")

        row["Люмн. Цена"] = 0
        row["Люмн. Цена со скидкой"] = 0
        row["Люмн. Остаток"] = 0

    return result


def process_order(order, file_price=None, price=None):
    if file_price:
        price = get_price(file_price)

    result = search(price, order)
    formalized_result = formalize_search_result(result)

    return formalized_result


# order = get_order()
# converted_order = convert_order(order)
#
# price = get_price()
# result = search(price, converted_order)
# formalized_result = formalize_search_result(result)
#
# # df = pd.DataFrame(formalized_result)
# # df.to_excel("Interresult.xls")