import pandas as pd
from buyer_order_converter import get_order, convert_order


def get_price(file_name="C:\\Users\\User1\\Downloads\\Lyumna_09_08.xls"):
    price = pd.read_excel(file_name, header=3, usecols=list(range(2, 13)))
    return price


def contains_cls_and_part(cls, part, name):
    if cls and part:
        return (f'ч.{part}' in name or  f'Часть {part}' in name)\
                and f'{cls} кл' in name
    elif cls:
        return f'{cls} кл' in name
    elif part:
         return f'ч.{part}' in name or  f'Часть {part}' in name
    else:
        return True


def arrange_key_words(key_words):
    reductions = ["раб", "тет", "учеб"]
    arranged_key_words = []
    if "раб" in key_words and "тет" in key_words:
        arranged_key_words.append("р/тет")
    if "учеб" in key_words:
        arranged_key_words.append("учб")
    for word in key_words:
        if word not in reductions:
            arranged_key_words.append(word)
    return arranged_key_words


def contains_key_words(name, order_row):
    key_words = arrange_key_words(order_row.get("key_words"))
    info = arrange_key_words(order_row.get("info"))
    cls = order_row.get("cls")
    part = order_row.get("part")
    if all([word in name for word in key_words]) \
            and order_row.get("name")[:-2] in name \
            and all([word in name for word in info]) \
            and order_row.get("author")[:-2] in name \
            and order_row.get("coauthor")[:-2] in name \
            and order_row.get("publish")[:-2] in name:
        return contains_cls_and_part(cls, part, name)
    else:
        return False


def delete_clones(lst):
    used_dicts = []

    for dct in lst:
        if dct in used_dicts:
            continue
        else:
            used_dicts.append(dct)

    return used_dicts


def search(price: pd.DataFrame, formalized_order, minimal_year):
    result = []

    for order_row in formalized_order:
        for price_row in price.iterrows():
            price_data = price_row[1]
            if not pd.isnull(price_data["Наименование"]):
                name = price_data["Наименование"].lower()
                year = int(price_data["Год изд."])
                if year >= minimal_year:
                    if contains_key_words(name, order_row):
                        print(order_row)
                        price_data = price_data.to_dict()
                        price_data["Количество"] = order_row.get("amount")
                        result.append(price_data)
            else:
                continue

    result = delete_clones(result)

    return result


def formalize_search_result(result):
    columns_for_del = ["Стандарт", "Заказ (количество)", "Unnamed: 9", "Товар в пути"]

    for row in result:
        for column in columns_for_del:
            del row[column]

        row["Год издания"] = row.pop("Год изд.")
        row["Люмн. Цена со скидкой"] = row.pop("Цена со скидкой")
        row["Люмн. Цена"] = row.pop("Цена")
        row["Люмн. Остаток"] = row.pop("Остаток")
        row["Люмн. Код"] = row.pop("Код")
        row["ISBN"] = int(row.pop("EAN"))

        print(row)

    return result


def process_order(order, file_price=None, price=None):
    if file_price:
        price = get_price(file_price)

    result = search(price, order, 2019)
    formalized_result = formalize_search_result(result)

    return formalized_result


# price = get_price()
# order = get_order()
# converted_order = convert_order(order)
#
# result = search(price, converted_order, 2019)
# # df = pd.DataFrame(formalize_search_result(result))
# # df.to_excel("tryResult.xls")
# # print(len(result))