import openpyxl
import requests
import time
import concurrent.futures


def download_files(name, link):
    try:
        print("downloading {}".format(name))
        data = requests.get(link).content
        file_name = f'{name}.zip'
        with open(file_name, 'wb') as _file:
            _file.write(data)
            return f'{name} was downloaded...'
    except Exception as e:
        print(e)

path = "E:\Codes\Python\Web-Automate-Selenium\Internshala-Python-Users.xlsx"
book = openpyxl.load_workbook(path)
sheet = book.get_sheet_by_name("contest_users")

names = [name.value for name in sheet["A"][1:]]
links = [link.value for link in sheet["D"][1:]]

with concurrent.futures.ThreadPoolExecutor() as executor:
	results = executor.map(download_files , names,links)
for result in results:
    print(result)