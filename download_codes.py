import openpyxl
import requests
import time
from zipfile import ZipFile
import concurrent.futures
import logging


logging.basicConfig(filename="logs.log", 
                    format='%(asctime)s %(message)s', 
                    filemode='w') 
logger=logging.getLogger()
logger.setLevel(logging.DEBUG) 


def unzip_files(file_name, date, name):
    '''Unzip and save the downloaded files sorted by date'''

    try:
        # opening the zip file in READ mode 
        with ZipFile(file_name, 'r') as zip:    
            zip.extractall(path = f'Downloads\\Contest\\{date}\\{name}//')   # extracting all the files 

    except Exception as e:
        # print(f"{name} not valid file")
        logger.warn(f"{name} not valid zip file") 


def download_files(name, date, link):
    '''Dowload and save zip files to a folder from given link'''

    try:
        print("downloading {}".format(name))
        data = requests.get(link).content
        file_name = f'Downloads\\Contest\\Zip\\{name}.zip'
        with open(file_name, 'wb') as _file:
            _file.write(data)
            
        unzip_files(file_name, date, name)
        return f'{name} was downloaded and unzipped...'
    except Exception as e:
        logger.warn(f"{name} not downloadable") 

path = "E:\Codes\Python\Web-Automate-Selenium\Internshala-Python-Users.xlsx"  #XL sheet path
book = openpyxl.load_workbook(path)
sheet = book.get_sheet_by_name("contest_users")   #fetch all user data

names = [name.value for name in sheet["A"][1:]]
dates = [date.value.split(" ")[0] for date in sheet["C"][1:]]
links = [link.value for link in sheet["D"][1:]]


with concurrent.futures.ThreadPoolExecutor() as executor:   #Concurrently download all files usuing multithreading
	results = executor.map(download_files , names, dates, links)
count = 0
for result in results:
    count +=1
    print(result,"  ", count)
