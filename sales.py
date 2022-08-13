import pandas as pd

def get_sales_data(file_name="C:\\Users\\User1\\Downloads\\Sistema_skidok_2022.xls"):
    sheet = pd.read_excel(io=file_name,  header=0, )
    for row in sheet.iterrows():
        print(row[1])


# get_sales_data()