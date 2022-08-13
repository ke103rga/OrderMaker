import pandas as pd

import lumna_searcher
from buyer_order_converter import get_order, convert_order
import interservice_searcher
import lumna


def merge(inter_result, lumna_result):
    for inter_row in inter_result:
        isbn = "".join(inter_row["ISBN"].split("-"))
        for lumna_row in lumna_result:
            if int(isbn) == int(lumna_row.get("ISBN")):
                inter_row["Люмн. Код"] = lumna_row["Люмн. Код"]
                inter_row["Люмн. Остаток"] = lumna_row["Люмн. Остаток"]
                inter_row["Люмн. Цена"] = lumna_row["Люмн. Цена"]
                inter_row["Люмн. Цена со скидкой"] = lumna_row["Люмн. Цена со скидкой"]

                lumna_result.remove(lumna_row)

    for lumna_row in lumna_result:
        lumna_row["Заказ"] = ""
        lumna_row["Заказ (количество)"] = ""
        lumna_row["Проверка"] = ""

        lumna_row["Люмн. Код"] = lumna_row["Люмн. Код"]
        lumna_row["Код"] = 0
        lumna_row["Интер. Цена"] = 0
        lumna_row["Интер. Цена со скидкой"] = 0
        lumna_row["Интер. Остаток"] = 0

    inter_result.extend(lumna_result)

    return inter_result


# order = get_order()
# converted_order = convert_order(order)
# lumna_price = lumna_searcher.get_price("C:\\Users\\User1\\Downloads\\Lyumna_09_08.xls")
# inter_price = interservice_searcher.get_price("")
#
#
# inter_result = interservice_searcher.process_order(file_price="C:\\Users\\User1\\Downloads\\knigi-polnyy-prays_3.xlsx", order=converted_order)
# inter_result = lumna_searcher.search(lumna_price, inter_result)
# lumna_result = lumna.process_order(price=lumna_price, order=converted_order)
#
#
# df = pd.DataFrame(inter_result)
# df.to_excel("Result.xls")