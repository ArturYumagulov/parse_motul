from module_1 import excel_reader, get_link, write_xls


lst = []
xls_path = "C:\\Users\\YumagulovA\\Desktop\\ozon\\motul.xlsx"
result_path = "C:\\Users\\YumagulovA\\Desktop\\ozon\\result.xlsx"

print("Скачивание....", end="")
for i in excel_reader(xls_path):
    correct_link = get_link(f"https://motul.store/search/?q={i}")
    print(correct_link)
    if correct_link is not None:
        print(".", end="")
        lst.append(correct_link)
print("\n", "Идет запись данных в файл")
write_xls(lst, result_path)
print("Записано успешно")
