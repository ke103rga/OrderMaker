import pandas as pd
import re


def get_order(file_name="C:\\Users\\User1\\Desktop\\OrderMakerDocuments\\Orders\\Invoice2.xls"):
    sheet = pd.read_excel(io=file_name, usecols=list(range(2, 40)), header=8)
    invoice = []

    fsm = {"name": "amount", "amount": "price", "price": "sum", "sum": "end"}

    for row in sheet.iterrows():
        series = row[1]
        state = "name"

        row_data = {"name": "",
                    "amount": "",
                    "price": "",
                    "sum": ""}

        for i in range(len(series)):
            if not pd.isnull(series[i]):
                if series[i] == "шт":
                    continue
                row_data[state] = series[i]
                state = fsm.get(state)
                if state == "end":
                    invoice.append(row_data)
                    break

    order = pd.DataFrame(invoice)
    # order.to_excel("NewInvoice.xls")
    return order


def convert_order(order: pd.DataFrame):
    converted_order = []
    for row in order.iterrows():
        data = row[1]
        order_data = {"key_words": "", "part": "", "author": ""}

        params = data["name"].split(". ")
        order_data["key_words"] = parse_key_words(params[0])
        order_data["part"] = find_part(params)
        order_data["author"] = params[0].split()[0].lower()
        order_data.update(parse_subject(params[1]))
        order_data["cls"] = find_cls(params)
        order_data.update(parse_book_info(params[-1]))

        for index, param in enumerate(params):
            if index not in [0, 1, len(params) - 1]:
                order_data["info"] += f" {param} "
        order_data["info"] = parse_key_words(order_data["info"].lower())

        order_data["amount"] = data["amount"]
        # order_data["price"] = data["price"]
        # order_data["sum"] = data["sum"]

        converted_order.append(order_data)

    return converted_order


def parse_subject(subject):
    sub = {"name": "", "cls": ""}

    class_patern = r"\dкл"
    cls = re.search(class_patern, subject)

    if cls is not None:
        cls = cls.span()
        sub["name"] = subject[:cls[0]].split()[0].lower()
        sub["cls"] = subject[cls[0]:cls[0] + 1]
    else:
        sub["name"] = subject.strip().split()[0].lower()
        sub["cls"] = ""

    return sub


def parse_book_info(book_info):
    info = {"program": "",
            "publish": "",
            "coauthor": "",
            "info": ""}

    patterns = {"program_pattern":  r'"[А-Яа-я ]+"',
                "publish_pattern":  r'\([А-Яа-я ]+\)',
                "coauthor_pattern":  r'/[А-Яа-я ]+/'}

    for pattern_name, pattern in patterns.items():
        if re.search(pattern, book_info):
            span = re.search(pattern, book_info).span()
            info[pattern_name.split("_")[0]] = book_info[span[0]+1:span[1]-1].lower()
            book_info = book_info.replace(book_info[span[0]:span[1]], "")

    info["info"] = " ".join(book_info.split()).lower()

    return info


def parse_key_words(key_words):
    keys = ["раб", "тет", "конт", "карт", "учеб", "атл", "проп"]
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


def find_part(params):
    part_pattern = r"Часть \d"
    for param in params:
        if re.search(part_pattern, param):
            span = re.search(r"\d", param).span()
            return param[span[0]:span[1]]
    return ""


def find_cls(params):
    part_pattern = r"\dкл"
    for param in params:
        if re.search(part_pattern, param):
            span = re.search(r"\d", param).span()
            return param[span[0]:span[1]]
    return ""


# order = get_order()
# [print(order_row) for order_row in convert_order(order)]


