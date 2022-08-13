import tkinter as tk
import tkinter.filedialog as fd
from tkinter import messagebox
from buyer_order_converter import get_order, convert_order
from xls_maker import create_xls_fie, arrange_xls_file
from merger import merge
import lumna_searcher
import lumna
import interservice_searcher
import pandas as pd


# def create_xls_fie(lst, filename="Order.xls"):
#     df = pd.DataFrame(lst)
#
#     df = df.reindex(columns=["Код", "Люмн. Код", "Наименование",
#                              "Артикул", "ISBN", "Автор",
#                              "Издательство", "Год издания", "Количество",
#                              "Остаток", "Интер. Цена", "Интер. Цена со скидкой", "Заказ",
#                              "Люмн. Остаток", "Люмн. Цена", "Люмн. Цена со скидкой", "Заказ (количество)",
#                              "Проверка"])
#
#     df.to_excel(filename)


def create_dir_path(file_path):
    rev_path = file_path[::-1]
    rev_path = rev_path[rev_path.index("/")+1:]
    return rev_path[::-1]


def choose_initial_dir(files_data):
    if files_data["interservis_price"]:
        initialdir = create_dir_path(files_data["interservis_price"])
    elif files_data["lumna_price"]:
        initialdir = create_dir_path(files_data["lumna_price"])
    elif files_data["invoice"]:
        initialdir = create_dir_path(files_data["invoice"])
    elif files_data["result_dir"]:
        initialdir = files_data["result_dir"]
    else:
        initialdir = "/"

    return initialdir


def choose_file(dict, key):
    filetypes = (("Excel файл", "*.xlsx *xls"), )

    initialdir = choose_initial_dir(dict)

    filename = fd.askopenfilename(title="Открыть файл", initialdir=initialdir,
                                  filetypes=filetypes)
    if filename:
        dict[key] = filename


def choose_directory(dict, key):
    initialdir = choose_initial_dir(dict)

    directory = fd.askdirectory(title="Открыть папку", initialdir=initialdir)
    if directory:
        dict[key] = directory


def process_order(files_data, filename, window):
    try:
        order = get_order(files_data["invoice"])
        converted_order = convert_order(order)

        int_price = interservice_searcher.get_price(files_data["interservis_price"])
        lumn_price = lumna_searcher.get_price(files_data["lumna_price"])

        result = interservice_searcher.search(int_price, converted_order)
        formalized_result = interservice_searcher.formalize_search_result(result)
        extend_result = lumna_searcher.search(lumn_price, formalized_result)

        lumna_result = lumna.process_order(order=converted_order, price=lumn_price)

        final_order = merge(extend_result, lumna_result)

        result_file_path = f'{files_data["result_dir"]}/{filename.get()}.xlsx'
        create_xls_fie(final_order, result_file_path)
        arrange_xls_file(result_file_path)

        messagebox.showinfo("Готово!", "Файл успешно сохранен в указаннной вами папке.")
    except Exception as ex:
        print(ex)
        messagebox.showinfo("Ошибка!", "Убедитесь в том что верно указали путь ко всем файлам.")
    finally:
        window.destroy()


def show_result():
    window = tk.Tk()
    window.title("Составитель заказов")
    window.geometry('600x500')
    window.resizable(False, False)

    files_data = {"interservis_price": "",
                  "lumna_price": "",
                  "invoice": "",
                  "result_dir": ""}

    search_int_file = "Выберите прайс Интерсервиса"
    lbl = tk.Label(window, text=search_int_file, font=('Times', 15))
    lbl.place(x=80, y=50)
    btn_file_int = tk.Button(text="Выбрать файл",
                         command=lambda:choose_file(files_data, "interservis_price"))
    btn_file_int.place(x=450, y=50)

    search_lumn_file = "Выберите прайс Люмны"
    lbl = tk.Label(window, text=search_lumn_file, font=('Times', 15))
    lbl.place(x=80, y=120)
    btn_file_int = tk.Button(text="Выбрать файл",
                             command=lambda: choose_file(files_data, "lumna_price"))
    btn_file_int.place(x=450, y=120)

    search_inv_file = "Выберите файл с накладной"
    lbl = tk.Label(window, text=search_inv_file, font=('Times', 15))
    lbl.place(x=80, y=190)
    btn_file_int = tk.Button(text="Выбрать файл",
                             command=lambda: choose_file(files_data, "invoice"))
    btn_file_int.place(x=450, y=190)

    choose_res_dir = "Выберите папку,\n в которую будет сохранен результат"
    lbl = tk.Label(window, text=choose_res_dir, font=('Times', 15))
    lbl.place(x=80, y=260)
    btn_file_int = tk.Button(text="Выбрать папку",
                             command=lambda: choose_directory(files_data, "result_dir"))
    btn_file_int.place(x=450, y=260)

    choose_file_name = "Введите название итогового файла"
    lbl = tk.Label(window, text=choose_file_name, font=('Times', 15))
    lbl.place(x=80, y=330)
    filename = tk.StringVar()
    entry = tk.Entry(width=10, textvariable=filename)
    entry.place(x=450, y=330)

    btn_process_order = tk.Button(text="Обработать заказ",
                             command=lambda: process_order(files_data, filename, window))
    btn_process_order.place(x=240, y=400)

    window.mainloop()

