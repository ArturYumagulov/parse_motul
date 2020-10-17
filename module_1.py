import openpyxl
import requests
from bs4 import BeautifulSoup as bs


def write_xls(lst):
    wb = openpyxl.Workbook()
    wb.create_sheet(title="Лист", index=0)
    sheet = wb["Лист"]
    for i in range(len(lst)):
        for j in lst[i]:
            value = str(j)
            cell = sheet.cell(row=i + 1, column=(lst[i].index(j)) + 1)
            cell.value = value
    wb.save('/home/zico/Desktop/example.xlsx')


def excel_reader(file_obj):
    result = []
    wb = openpyxl.load_workbook(file_obj)
    sheet = wb['МП']
    for row in sheet['A']:
        result.append(str(row.value))
    return result


def get_link(url):
    r = requests.get(url)
    if r.status_code == 200:
        soup = bs(r.text, "html.parser")
        divs = soup.find_all('div', attrs={"class": "catalog_content_products-item"})
        for div in divs:
            types = div.find('span', attrs={'class': 'catalog_content_products-category'}).text
            if types in ["Моторные масла Motul", "Трансмиссионные масла Motul", "Моторные масла 300V Motul"]:
                product = div.find('img', attrs={"class": 'catalog_content_products-img'})
                litres = div.find('span', attrs={'class': 'catalog-card-capacity__liter active-liter'})
                if "Трансмиссионные масла" in types:
                    types = "Автохимия.Трансмиссионное"
                elif "Моторные масла" in types:
                    types = "Автохимия. Масло моторное"
                if litres is not None:
                    price = litres.get('data-price')
                    articul = litres.get('data-articul')
                image = product.get("src")
                weight = 0
                length = 0
                height = 0
                if "5" in litres.text.strip():
                    litres = 5000
                    weight = 150
                    length = 200
                    height = 300
                elif "4" in litres.text.strip():
                    litres = 4000
                    weight = 150
                    length = 200
                    height = 300
                elif "2" in litres.text.strip():
                    litres = 2000
                    weight = 10
                    length = 10
                    height = 20
                elif "1" in litres.text.strip():
                    litres = 1000
                    weight = 10
                    length = 10
                    height = 20
                if len(image) != 0:
                    image_list = (image.split('/'))
                    correct_url = f"https://motul.store/{image_list[1]}/{image_list[2]}/{image_list[3]}/{image_list[4]}/" \
                                  f"700_700_1/{image_list[-1]}"
                    print([articul, int(price.replace(" руб.", "").replace(' ', '')), types, litres, weight,
                            length, height, correct_url])
                    return [articul, int(price.replace(" руб.", "").replace(' ', '')), types, litres, weight,
                            length, height, correct_url]
                else:
                    return None
            else:
                continue
    else:
        print(r.status_code)


count = 1
lst = []
xls_path = "/home/zico/Desktop/motul.xlsx"
with open("/home/zico/Desktop/file.txt", 'a') as f:
    for i in excel_reader(xls_path):
        correct_link = get_link(f"https://motul.store/search/?q={i}")
        if correct_link is not None:
            lst.append(correct_link)



# x = get_link(f"https://motul.store/search/?q={100198}")
# print(len(x))
# print(x)

